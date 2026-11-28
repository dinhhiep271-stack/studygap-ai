"""Kiểm thử cho ``POST /process-book`` và ``GET /process-book/{job_id}``.

Không gọi DeepSeek thật: client được thay bằng ``FakeBookClient`` nên test chạy
offline. PDF nhiều trang được dựng trực tiếp trong file này.

Các nhóm kiểm thử:
    run_split_tests    chia chương theo mục lục / tiêu đề chương / chia đều trang
    run_limit_tests    giới hạn số chương, cờ ``background``, gộp phần dư
    run_job_store_tests  vòng đời job trong RAM (tạo/cập nhật/TTL/giới hạn số job)
    asyncio.run(run_route_tests())            route đồng bộ (background=false)
    asyncio.run(run_background_job_tests()) job ngầm: 202 + BackgroundTasks + poll
    run_http_tests     tầng HTTP multipart qua ``TestClient``

Chạy:
    .\\venv\\Scripts\\python.exe test_process_book.py

Exit code 0 = tất cả pass, 1 = có kiểm thử fail.
"""

from __future__ import annotations

import asyncio
import datetime as dt
import io
import json
import re
import sys
from types import SimpleNamespace

from fastapi import BackgroundTasks, HTTPException, Response
from starlette.datastructures import Headers, UploadFile

import main


# Windows console mặc định không phải UTF-8 -> ép UTF-8 để in được tiếng Việt.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


_failures: list[str] = []


def ok(cond: bool, label: str) -> None:
    """Ghi nhận một kiểm thử; không dừng ngay để thấy hết lỗi trong một lần chạy."""
    print(("PASS  " if cond else "FAIL  ") + label)
    if not cond:
        _failures.append(label)


def make_upload(
    filename: str, data: bytes, content_type: str = "application/pdf"
) -> UploadFile:
    """Tạo ``UploadFile`` trong bộ nhớ giống những gì FastAPI nhận từ multipart."""
    return UploadFile(
        file=io.BytesIO(data),
        filename=filename,
        headers=Headers({"content-type": content_type}),
    )


def build_multipage_pdf(page_texts: list[str]) -> bytes:
    """Dựng PDF nhiều trang, mỗi trang một dòng text mà pypdf đọc lại được.

    Object: 1 catalog, 2 pages, 3 font; sau đó mỗi trang là 2 object liên tiếp
    (page = 4 + 2i, content = 5 + 2i).
    """
    contents: list[bytes] = []
    for text in page_texts:
        escaped = text.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
        line = f"BT /F1 10 Tf 40 740 Td ({escaped}) Tj ET"
        contents.append(line.encode("latin-1", "replace"))

    count = len(page_texts)
    page_objs = [4 + 2 * index for index in range(count)]
    content_objs = [5 + 2 * index for index in range(count)]
    kids = b" ".join(f"{number} 0 R".encode() for number in page_objs)

    objs: list[bytes] = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [" + kids + f"] /Count {count} >>".encode(),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    for index in range(count):
        objs.append(
            (
                f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
                f"/Contents {content_objs[index]} 0 R "
                "/Resources << /Font << /F1 3 0 R >> >> >>"
            ).encode()
        )
        body = contents[index]
        objs.append(
            b"<< /Length "
            + str(len(body)).encode()
            + b" >>\nstream\n"
            + body
            + b"\nendstream"
        )

    out = bytearray(b"%PDF-1.4\n")
    offsets: list[int] = []
    for number, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += f"{number} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(objs) + 1}\n".encode() + b"0000000000 65535 f \n"
    for offset in offsets:
        out += f"{offset:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\n"
        f"startxref\n{xref_pos}\n%%EOF\n"
    ).encode()
    return bytes(out)


def toc_book_pages() -> list[str]:
    """Sách 8 trang có mục lục: 3 chương in ở trang 3, 5, 7 (kèm mục con 2.1)."""
    return [
        "Inhaltsverzeichnis",
        (
            "1 Grundlagen der BWL 3\n"
            "2 Recht und Vertraege 5\n"
            "2.1 Vertragsarten 6\n"
            "3 Marketing 7"
        ),
        "1 Grundlagen der BWL\nMARKER-ch1\nBetriebswirtschaft und Kennzahlen",
        "Fortsetzung Kapitel 1\nMARKER-ch1\nWirtschaftlichkeit",
        "2 Recht und Vertraege\nMARKER-ch2\nBGB und HGB",
        "Fortsetzung Kapitel 2\nMARKER-ch2\nHandelsregister",
        "3 Marketing\nMARKER-ch3\nMarketing-Mix",
        "Anhang\nQuellenverzeichnis",
    ]


_PLAN_ENTRY_RE = re.compile(r"([a-z_]+)=(\d+)")


def parse_plan(prompt: str) -> list[str]:
    """Đọc dòng "- Số câu theo từng dạng: ..." trong prompt thành plan dạng câu."""
    plan: list[str] = []
    for line in prompt.splitlines():
        if line.startswith("- Số câu theo từng dạng:"):
            for qtype, amount in _PLAN_ENTRY_RE.findall(line):
                plan.extend([qtype] * int(amount))
    return plan


def make_question(qtype: str, index: int, marker: str) -> dict:
    """Một câu hỏi hợp lệ theo hợp đồng JSON (đủ dùng cho fake client)."""
    tag = f"{marker} #{index}"
    if qtype == "multiple_response":
        return {
            "id": index,
            "type": qtype,
            "question": f"Chọn nhiều đáp án {tag}",
            "options": {"A": "Đúng 1", "B": "Sai", "C": "Đúng 2", "D": "Sai khác"},
            "answer": ["A", "C"],
            "explanation": "A và C đúng",
            "difficulty": "Trung bình",
        }
    if qtype == "calculation":
        return {
            "id": index,
            "type": qtype,
            "question": f"Tính toán {tag}",
            "given": ["Doanh thu 100", "Chi phí 60"],
            "formula": "Lợi nhuận = Doanh thu - Chi phí",
            "answer": "40",
            "explanation": "100 - 60 = 40",
            "difficulty": "Khó",
        }
    if qtype == "fill_in_the_blank":
        return {
            "id": index,
            "type": qtype,
            "question": f"Điền khuyết {tag}: Lợi nhuận = Doanh thu - ___",
            "answer": ["Chi phí"],
            "explanation": "Công thức lợi nhuận",
            "difficulty": "Dễ",
        }
    if qtype == "table":
        return {
            "id": index,
            "type": "table",
            "question": f"Điền bảng {tag}",
            "table_data": {
                "headers": ["Chỉ tiêu", "Số tiền"],
                "rows": [["Doanh thu", "100"], ["Chi phí", "___"]],
            },
            "answer": [{"row": 1, "column": 1, "value": "60"}],
            "explanation": "Chi phí = 100 - 40",
            "difficulty": "Trung bình",
        }
    return {
        "id": index,
        "type": "multiple_choice",
        "question": f"Trắc nghiệm {tag}",
        "options": {"A": "Đúng", "B": "Sai", "C": "Sai", "D": "Sai"},
        "answer": "A",
        "explanation": "A đúng",
        "difficulty": "Dễ",
    }


class FakeBookClient:
    """DeepSeek giả: mỗi lần gọi trả JSON cho ĐÚNG chương có trong prompt.

    Marker ``MARKER-*`` lấy từ nội dung chương nên test khẳng định được kết quả
    của từng chương không trộn nội dung chương khác. ``fail_markers`` mô phỏng
    DeepSeek lỗi ở một chương để kiểm thử cô lập lỗi.
    """

    def __init__(self, *, fail_markers: set[str] | None = None) -> None:
        self.calls: list[dict] = []
        self.fail_markers = fail_markers or set()
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs) -> SimpleNamespace:
        self.calls.append(kwargs)
        content = kwargs["messages"][-1]["content"]
        marker = next(
            (
                line.strip()
                for line in content.splitlines()
                if line.strip().startswith("MARKER-")
            ),
            "UNKNOWN",
        )
        if marker in self.fail_markers:
            raise RuntimeError(f"DeepSeek giả lỗi cho {marker}")
        plan = parse_plan(kwargs["messages"][0]["content"])
        payload = {
            "title": f"Chương {marker}",
            "teaching": {
                "summary_de": f"Zusammenfassung {marker}",
                "summary_vi": marker,
                "concepts_de": [
                    {
                        "name": f"Begriff {marker}",
                        "explanation": "Deutsche Erklärung",
                        "sources": ["buch.pdf"],
                    }
                ],
                "concepts_vi": [
                    {
                        "name": f"Khái niệm {marker}",
                        "explanation": "Giải thích tiếng Việt",
                        "sources": ["buch.pdf"],
                    }
                ],
            },
            "questions": [
                make_question(qtype, index, marker)
                for index, qtype in enumerate(plan, start=1)
            ],
        }
        return SimpleNamespace(
            choices=[
                SimpleNamespace(
                    message=SimpleNamespace(
                        content=json.dumps(payload, ensure_ascii=False)
                    )
                )
            ],
            model="deepseek-chat",
            usage=SimpleNamespace(prompt_tokens=5, completion_tokens=7, total_tokens=12),
        )


def run_split_tests() -> None:
    """Chia sách: mục lục, ghép số trang in -> trang PDF, tiêu đề chương, chia đều."""
    # --- dòng mục lục ----------------------------------------------------------
    ok(
        main._split_toc_line("Kapitel 4: Kostenrechnung ....... 42", 60)
        == ("4", "Kostenrechnung", 42),
        "mục lục: 'Kapitel 4: Kostenrechnung .... 42' -> ('4', 'Kostenrechnung', 42)",
    )
    ok(
        main._split_toc_line("1 Einleitung 3", 8) == ("1", "Einleitung", 3),
        "mục lục: '1 Einleitung 3' -> ('1', 'Einleitung', 3)",
    )
    ok(
        main._split_toc_line("1 Einleitung", 8) is None,
        "mục lục: dòng không có số trang bị bỏ",
    )
    sub_entry = main._split_toc_line("2.1 Vertragsarten ... 6", 8)
    ok(
        sub_entry is None or sub_entry[0] == "2",
        "mục lục: mục con '2.1' không tạo chương cấp 1 riêng",
    )
    ok(
        main._split_toc_line("12 Anhang 999", 8) is None,
        "mục lục: số trang vô lý (999 > số trang) bị bỏ",
    )
    ok(
        main._split_toc_line("Inhaltsverzeichnis", 8) is None,
        "mục lục: tiêu đề trang bị bỏ",
    )

    # --- tìm mục lục + ghép trang ---------------------------------------------
    pages = toc_book_pages()
    split = main.split_book(pages)
    labels = [chapter.label for chapter in split.chapters]
    ok(split.method == "toc", f"chia theo mục lục (method={split.method})")
    ok(split.toc_pages == (1, 2), f"mục lục nằm trang 1-2 (nhận {split.toc_pages})")
    ok(
        labels == ["1 Grundlagen der BWL", "2 Recht und Vertraege", "3 Marketing"],
        f"bỏ mục con 2.1, còn 3 chương cấp 1 (nhận {labels})",
    )
    ok(
        [chapter.printed_page for chapter in split.chapters] == [3, 5, 7],
        "giữ đúng số trang in của mục lục",
    )
    ok(
        [chapter.start_page for chapter in split.chapters] == [3, 5, 7],
        "số trang in khớp trang PDF -> bắt đầu ở trang 3, 5, 7",
    )
    ok(
        [chapter.end_page for chapter in split.chapters] == [4, 6, 8],
        "trang kết thúc nối tiếp tới hết sách (4, 6, 8)",
    )

    # --- số trang in lệch trang PDF (bìa + trang phụ) --------------------------
    shifted = [
        "Inhaltsverzeichnis\n1 Einleitung 3\n2 Hauptteil 8",
        "Vorwort",
        "Impressum",
        "Einleitung zum Buch",
        "1 Einleitung\nMARKER-ein",
        "Text nach Einleitung",
        "Noch Text",
        "Zwischenseite",
        "Hauptteil\nMARKER-haupt",
        "Fortsetzung Hauptteil",
        "Anhang",
    ]
    shifted_split = main.split_book(shifted)
    ok(shifted_split.method == "toc", "sách có trang phụ vẫn chia được theo mục lục")
    ok(
        [chapter.start_page for chapter in shifted_split.chapters] == [5, 10],
        "bù offset bìa/trang phụ: trang in 3 -> trang PDF 5, trang in 8 -> trang PDF 10",
    )
    ok(
        [chapter.end_page for chapter in shifted_split.chapters] == [9, 11],
        "khoảng trang PDF sau khi bù offset là 5-9 và 10-11",
    )

    # --- không có mục lục: dò tiêu đề chương ----------------------------------
    heading_pages = [
        "Titelblatt\nStudygap AI",
        "Inhaltsverzeichnis",
        "Kapitel 1: Grundlagen der BWL\nText",
        "Vertiefung ohne Nummer",
        "Kapitel 2: Recht\nText",
    ]
    heading_split = main.split_book(heading_pages)
    ok(
        heading_split.method == "headings",
        f"không có mục lục -> dò tiêu đề chương (method={heading_split.method})",
    )
    ok(
        [(c.start_page, c.end_page) for c in heading_split.chapters] == [(3, 4), (5, 5)],
        "dò được 'Kapitel 1' ở trang 3 và 'Kapitel 2' ở trang 5",
    )
    ok(
        [c.label for c in heading_split.chapters]
        == ["1 Grundlagen der BWL", "2 Recht"],
        "nhãn chương lấy đúng số + tên in trên trang",
    )

    # --- không mục lục, không tiêu đề: chia đều theo trang ---------------------
    plain = [f"Nur Fliesstext auf Seite {index}." for index in range(1, 6)]
    original_chunk = main.BOOK_FALLBACK_CHAPTER_PAGES
    main.BOOK_FALLBACK_CHAPTER_PAGES = 2
    try:
        chunked = main.split_book(plain)
    finally:
        main.BOOK_FALLBACK_CHAPTER_PAGES = original_chunk
    ok(
        chunked.method == "pages",
        f"không có dấu hiệu chương -> chia đều theo trang (method={chunked.method})",
    )
    ok(
        [(c.start_page, c.end_page) for c in chunked.chapters] == [(1, 2), (3, 4), (5, 5)],
        "chia 5 trang thành 3 phần 1-2, 3-4, 5-5",
    )
    ok(
        chunked.chapters[0].start_page == 1
        and chunked.chapters[-1].end_page == len(plain),
        "chia đều phủ toàn bộ sách, không mất trang",
    )


def run_limit_tests() -> None:
    """Giới hạn số chương + cờ ``background`` của form."""
    chapters = [
        main.BookChapter(
            index=index,
            number=str(index),
            title=title,
            printed_page=index,
            start_page=start,
            end_page=end,
        )
        for index, (title, start, end) in enumerate(
            [("A", 1, 3), ("B", 4, 6), ("C", 7, 9)], start=1
        )
    ]

    kept, trimmed = main.limit_book_chapters(chapters, 2, 9)
    ok(len(kept) == 2 and trimmed, "quá số chương: giữ 2 chương và báo trimmed")
    ok(kept[-1].end_page == 9, "chương cuối được nới tới hết sách (không mất nội dung)")
    ok(
        kept[-1].title == "B (và phần còn lại)",
        f"chương cuối ghi rõ phần gộp (nhận {kept[-1].title!r})",
    )
    ok(chapters[2].end_page == 9 and chapters[2].title == "C", "không sửa chương gốc")

    same, trimmed_same = main.limit_book_chapters(chapters, 5, 9)
    ok(len(same) == 3 and not trimmed_same, "limit lớn hơn số chương: giữ nguyên")
    untouched, _ = main.limit_book_chapters(chapters, 0, 9)
    ok(len(untouched) == 3, "limit=0 (không giới hạn) trả đủ chương")

    ok(
        main.clamp_book_chapter_limit(0) == main.MAX_BOOK_CHAPTERS
        and main.clamp_book_chapter_limit(-3) == main.MAX_BOOK_CHAPTERS,
        "max_chapters=0/âm -> mặc định MAX_BOOK_CHAPTERS",
    )
    ok(
        main.clamp_book_chapter_limit(10_000) == main.MAX_BOOK_CHAPTERS
        and main.clamp_book_chapter_limit(3) == 3,
        "max_chapters được kẹp trong 1..MAX_BOOK_CHAPTERS",
    )

    ok(
        main.resolve_book_background("") is main.BOOK_BACKGROUND_BY_DEFAULT,
        "background rỗng -> mặc định (chạy ngầm)",
    )
    ok(
        not main.resolve_book_background("false")
        and not main.resolve_book_background("0")
        and not main.resolve_book_background("OFF"),
        "background=false/0/off -> chạy đồng bộ",
    )
    ok(main.resolve_book_background("true"), "background=true -> chạy ngầm")


def run_job_store_tests() -> None:
    """Vòng đời job trong bộ nhớ: tạo, cập nhật, đọc bản sao, TTL, giới hạn số job."""
    main.book_jobs.clear()
    chapter = main.BookChapter(
        index=1,
        number="1",
        title="Grundlagen",
        printed_page=1,
        start_page=1,
        end_page=2,
    )
    job = main.create_book_job(
        book={"filename": "buch.pdf", "pages": 2},
        chapters=[chapter],
        request={"lang": "vi", "count": 2},
    )
    ok(job["status"] == "queued", "job mới ở trạng thái queued")
    ok(
        job["progress"] == {"completed": 0, "total": 1, "failed": 0},
        "job mới có progress 0/1",
    )
    ok(
        job["chapters"][0]["start_page"] == 1 and job["results"] == [],
        "job mới có danh sách chương và chưa có kết quả",
    )
    ok(
        len(job["job_id"]) == 12 and job["job_id"] in main.book_jobs,
        "job_id 12 ký tự được lưu vào store",
    )

    copy = main.get_book_job(job["job_id"])
    assert copy is not None
    copy["status"] = "hacked"
    copy["results"].append({"success": True})
    fresh = main.get_book_job(job["job_id"])
    ok(
        fresh is not None and fresh["status"] == "queued" and fresh["results"] == [],
        "get_book_job trả bản sao (client không sửa được store)",
    )

    main.update_book_job(job["job_id"], status="running")
    ok(
        main.get_book_job(job["job_id"])["status"] == "running",
        "update_book_job đổi được trạng thái",
    )
    main.update_book_job("khong-ton-tai", status="running")
    ok(main.get_book_job("khong-ton-tai") is None, "cập nhật job lạ không lỗi, không tạo job")

    # Job hết TTL bị dọn khi có request đọc.
    stale = (
        dt.datetime.now(dt.timezone.utc)
        - dt.timedelta(seconds=main.BOOK_JOB_TTL_SECONDS + 5)
    ).isoformat(timespec="seconds")
    main.update_book_job(job["job_id"], status="completed", finished_at=stale)
    ok(main.get_book_job(job["job_id"]) is None, "job quá TTL bị xoá khỏi store")

    # Chỉ giữ tối đa MAX_BOOK_JOBS job.
    main.book_jobs.clear()
    original_max = main.MAX_BOOK_JOBS
    main.MAX_BOOK_JOBS = 2
    try:
        ids = [
            main.create_book_job(book={}, chapters=[chapter], request={})["job_id"]
            for _ in range(3)
        ]
        main.get_book_job(ids[2])  # prune chạy khi có request đọc
    finally:
        main.MAX_BOOK_JOBS = original_max
    ok(
        len(main.book_jobs) == 2 and ids[0] not in main.book_jobs,
        "quá MAX_BOOK_JOBS thì job cũ nhất bị xoá",
    )
    main.book_jobs.clear()


async def run_route_tests() -> None:
    """Gọi thẳng ``process_book`` (bỏ qua HTTP): đồng bộ, lỗi từng chương, job ngầm."""
    pdf = build_multipage_pdf(toc_book_pages())
    pages = main.extract_pdf_pages(pdf)
    ok(len(pages) == 8, f"PDF nhiều trang đọc lại đủ 8 trang (nhận {len(pages)})")
    ok("MARKER-ch1" in pages[2], "trích đúng text của trang 3")
    ok(pdf.startswith(b"%PDF-") and main._looks_like_pdf(pdf), "PDF dựng trong test hợp lệ")

    original_get_client = main.get_deepseek_client
    fake = FakeBookClient()
    main.get_deepseek_client = lambda: fake  # type: ignore[assignment]
    try:
        results = await main.process_book(
            BackgroundTasks(),
            Response(),
            make_upload("buch.pdf", pdf),
            lang="vi",
            subject="BWL",
            count=2,
            types="multiple_choice",
            background="false",
            max_chapters=0,
        )
        ok(isinstance(results, list), "background=false -> trả thẳng MẢNG kết quả")
        if not isinstance(results, list):
            return
        ok(len(results) == 3, f"xử lý đủ 3 chương (nhận {len(results)})")
        if len(results) != 3:
            return
        ok(all(item["success"] for item in results), "cả 3 chương đều thành công")
        ok(
            [item["chapter"]["label"] for item in results]
            == ["1 Grundlagen der BWL", "2 Recht und Vertraege", "3 Marketing"],
            "mỗi phần tử gắn đúng metadata chương",
        )
        ok(
            [item["chapter"]["start_page"] for item in results] == [3, 5, 7]
            and [item["chapter"]["end_page"] for item in results] == [4, 6, 8],
            "khoảng trang từng chương là 3-4, 5-6, 7-8",
        )
        ok(
            [item["chapter"]["pages"] for item in results] == [2, 2, 2],
            "số trang của mỗi chương đúng",
        )
        ok(
            [item["filename"] for item in results]
            == [
                "buch - 1 Grundlagen der BWL",
                "buch - 2 Recht und Vertraege",
                "buch - 3 Marketing",
            ],
            "tên tài liệu = tên sách + nhãn chương",
        )
        ok(
            [item["data"]["teaching"]["summary_vi"] for item in results]
            == ["MARKER-ch1", "MARKER-ch2", "MARKER-ch3"],
            "mỗi chương chỉ dùng nội dung của chính nó (không trộn chương khác)",
        )
        ok(
            [item["documents_used"] for item in results]
            == [
                ["buch - 1 Grundlagen der BWL"],
                ["buch - 2 Recht und Vertraege"],
                ["buch - 3 Marketing"],
            ],
            "mode per_file: prompt mỗi chương chỉ có tài liệu của chương đó",
        )
        ok(all(item["mode"] == "per_file" for item in results), "dùng chung pipeline per_file")
        ok(
            len(fake.calls) == 3,
            f"gọi DeepSeek đúng 3 lần (1 lần/chương), nhận {len(fake.calls)}",
        )
        ok(
            all(item["count"] == 2 and len(item["data"]["questions"]) == 2 for item in results),
            "mỗi chương có đúng 2 câu hỏi như yêu cầu",
        )
        ok(
            all(item["chapter"]["characters_extracted"] > 0 for item in results),
            "metadata chương có số ký tự trích được",
        )

        # --- một chương lỗi không làm hỏng cả sách ----------------------------
        failing = FakeBookClient(fail_markers={"MARKER-ch2"})
        main.get_deepseek_client = lambda: failing  # type: ignore[assignment]
        partial = await main.process_book(
            BackgroundTasks(),
            Response(),
            make_upload("buch.pdf", pdf),
            count=2,
            types="multiple_choice",
            background="false",
        )
        ok(
            isinstance(partial, list) and len(partial) == 3,
            "vẫn trả đủ 3 phần tử khi 1 chương lỗi",
        )
        if isinstance(partial, list) and len(partial) == 3:
            ok(
                [item["success"] for item in partial] == [True, False, True],
                "chỉ chương 2 thất bại, chương 1 và 3 vẫn xong",
            )
            ok(
                partial[1]["status_code"] == 500
                and "DeepSeek giả" in partial[1]["error"],
                "phần tử lỗi có status_code 500 + thông điệp lỗi",
            )
            ok(
                partial[1]["chapter"]["label"] == "2 Recht und Vertraege",
                "phần tử lỗi vẫn giữ metadata chương",
            )

        # --- giới hạn số chương: phần dư gộp vào chương cuối -------------------
        main.get_deepseek_client = lambda: FakeBookClient()  # type: ignore[assignment]
        single = await main.process_book(
            BackgroundTasks(),
            Response(),
            make_upload("buch.pdf", pdf),
            count=2,
            types="multiple_choice",
            background="false",
            max_chapters=1,
        )
        ok(
            isinstance(single, list) and len(single) == 1,
            "max_chapters=1 -> xử lý 1 chương",
        )
        if isinstance(single, list) and len(single) == 1:
            ok(single[0]["chapter"]["end_page"] == 8, "chương gộp phủ tới trang cuối sách")
            ok(
                "và phần còn lại" in single[0]["chapter"]["title"],
                "tên chương gộp ghi rõ phần còn lại",
            )

        # --- lỗi đầu vào ------------------------------------------------------
        for label, upload, expected in (
            (
                "file không phải PDF -> 415",
                make_upload("a.png", b"\x89PNG\r\n\x1a\n", "image/png"),
                415,
            ),
            ("file rỗng -> 400", make_upload("leer.pdf", b"", "application/pdf"), 400),
        ):
            try:
                await main.process_book(BackgroundTasks(), Response(), upload)
            except HTTPException as exc:
                ok(exc.status_code == expected, f"{label} (nhận {exc.status_code})")
            else:
                ok(False, f"{label} nhưng không raise HTTPException")
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]


async def run_background_job_tests() -> None:
    """``background=true``: trả 202 + job, chạy BackgroundTasks rồi poll ``GET``."""
    pdf = build_multipage_pdf(toc_book_pages())
    original_get_client = main.get_deepseek_client
    main.get_deepseek_client = lambda: FakeBookClient()  # type: ignore[assignment]
    main.book_jobs.clear()
    try:
        tasks = BackgroundTasks()
        response = Response()
        job = await main.process_book(
            tasks,
            response,
            make_upload("buch.pdf", pdf),
            count=2,
            types="multiple_choice",
            background="",
        )
        ok(
            response.status_code == 202,
            f"background mặc định -> 202 (nhận {response.status_code})",
        )
        ok(isinstance(job, dict) and job["status"] == "queued", "trả job queued ngay")
        if not isinstance(job, dict):
            return
        ok(
            job["progress"] == {"completed": 0, "total": 3, "failed": 0},
            "job mới có progress 0/3",
        )
        ok(
            job["book"]["split_method"] == "toc"
            and job["book"]["chapters_detected"] == 3
            and job["book"]["chapters_processed"] == 3,
            "job ghi cách chia + số chương dò được/xử lý",
        )
        ok(
            job["book"]["toc_pages"] == [1, 2] and job["book"]["pages"] == 8,
            "job ghi trang mục lục và tổng số trang PDF",
        )
        ok(
            len(job["chapters"]) == 3 and job["chapters"][0]["start_page"] == 3,
            "job kèm danh sách chương đã chia",
        )
        ok(
            job["request"]["lang"] == "vi"
            and job["request"]["count"] == 2
            and job["request"]["background"] is True,
            "job ghi lại yêu cầu của client",
        )
        ok(
            job["results"] == [] and job["finished_at"] is None,
            "job mới chưa có kết quả và chưa kết thúc",
        )
        ok(job["job_id"] in main.book_jobs, "job được lưu vào store")
        ok(len(tasks.tasks) == 1, "route đăng ký đúng 1 background task")

        for task in tasks.tasks:
            task.func(*task.args, **task.kwargs)

        polled = await main.get_process_book_job(job["job_id"])
        ok(
            polled["status"] == "completed",
            f"chạy background task -> completed (nhận {polled['status']})",
        )
        ok(
            polled["progress"] == {"completed": 3, "total": 3, "failed": 0},
            "progress cuối 3/3, 0 lỗi",
        )
        ok(
            [item["data"]["teaching"]["summary_vi"] for item in polled["results"]]
            == ["MARKER-ch1", "MARKER-ch2", "MARKER-ch3"],
            "poll được kết quả đúng nội dung từng chương",
        )
        ok(
            polled["finished_at"] is not None and polled["error"] is None,
            "job completed có finished_at và không có error",
        )

        # --- job hỏng: pipeline ném lỗi -> status failed ----------------------
        original_process = main.process_book_chapters
        main.process_book_chapters = lambda *args, **kwargs: (_ for _ in ()).throw(  # type: ignore[assignment]
            RuntimeError("boom")
        )
        try:
            broken = main.create_book_job(book={}, chapters=[], request={})
            main.run_book_job(
                broken["job_id"], [], [], book_stem="x", subject="", lang="vi", plan=[]
            )
        finally:
            main.process_book_chapters = original_process  # type: ignore[assignment]
        broken_job = main.get_book_job(broken["job_id"])
        ok(
            broken_job is not None
            and broken_job["status"] == "failed"
            and "boom" in (broken_job["error"] or ""),
            "job lỗi -> status failed kèm error",
        )

        try:
            await main.get_process_book_job("khong-ton-tai")
        except HTTPException as exc:
            ok(exc.status_code == 404, f"job không tồn tại -> 404 (nhận {exc.status_code})")
        else:
            ok(False, "job không tồn tại nhưng không raise 404")
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]
        main.book_jobs.clear()


def run_http_tests() -> None:
    """Multipart thật qua ASGI: 202 + job_id, poll GET, 404 job lạ, 415 file sai."""
    try:
        from fastapi.testclient import TestClient
    except Exception as exc:  # noqa: BLE001 - httpx là dependency tuỳ chọn của test
        print(f"SKIP  TestClient không dùng được ({exc}); bỏ qua các kiểm thử HTTP")
        return

    pdf = build_multipage_pdf(toc_book_pages())
    original_get_client = main.get_deepseek_client
    main.get_deepseek_client = lambda: FakeBookClient()  # type: ignore[assignment]
    main.book_jobs.clear()
    try:
        with TestClient(main.app) as client:
            response = client.post(
                "/process-book",
                files={"file": ("buch.pdf", pdf, "application/pdf")},
                data={
                    "count": "2",
                    "types": "multiple_choice",
                    "lang": "vi",
                    "subject": "BWL",
                },
            )
            ok(
                response.status_code == 202,
                f"HTTP: POST /process-book -> 202 (nhận {response.status_code})",
            )
            job = response.json()
            ok(
                job.get("status") in {"queued", "running", "completed"},
                "HTTP: body là job có status",
            )
            job_id = job.get("job_id", "")
            ok(bool(job_id) and job.get("progress", {}).get("total") == 3, "HTTP: job_id + tổng 3 chương")

            response = client.get(f"/process-book/{job_id}")
            ok(response.status_code == 200, "HTTP: GET /process-book/{job_id} -> 200")
            payload = response.json()
            ok(
                payload["status"] == "completed",
                f"HTTP: background task chạy xong (status={payload['status']})",
            )
            ok(
                len(payload["results"]) == 3
                and [item["chapter"]["start_page"] for item in payload["results"]]
                == [3, 5, 7],
                "HTTP: job có 3 kết quả chương đúng trang",
            )
            ok(
                payload["results"][0]["data"]["teaching"]["summary_vi"] == "MARKER-ch1",
                "HTTP: kết quả chương 1 đúng nội dung",
            )

            response = client.get("/process-book/khong-ton-tai")
            ok(response.status_code == 404, "HTTP: job không tồn tại -> 404")

            response = client.post(
                "/process-book",
                files={"file": ("a.png", b"\x89PNG\r\n\x1a\n", "image/png")},
            )
            ok(response.status_code == 415, "HTTP: file không phải PDF -> 415")

            response = client.post("/process-book")
            ok(response.status_code == 422, "HTTP: thiếu file -> 422")
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]
        main.book_jobs.clear()


if __name__ == "__main__":
    run_split_tests()
    run_limit_tests()
    run_job_store_tests()
    asyncio.run(run_route_tests())
    asyncio.run(run_background_job_tests())
    run_http_tests()

    if _failures:
        print(f"\n{len(_failures)} kiểm thử FAIL:")
        for label in _failures:
            print(f"  - {label}")
        sys.exit(1)
    print("\nALL CHECKS PASSED")
