"""Kiểm thử cho ``POST /process-document`` (batch, cross-document) và ``POST /generate-exam``.

Không gọi DeepSeek thật: client được thay bằng một đối tượng giả (``FakeClient``)
nên test chạy offline và không tốn quota. PDF mẫu được dựng trực tiếp trong file
này nên không cần fixture nhị phân.

Các nhóm kiểm thử:
    run_config_tests         cấu hình / bảo mật API key / loader ``.env``
    run_prompt_tests         prompt + hợp đồng JSON ``teaching`` / ``questions``
    run_contract_tests       helper chuẩn hoá, phân bổ dạng câu, kẹp số câu
    run_format_tests         đọc .pdf/.docx/.xlsx/.txt/.csv + lỗi định dạng
    run_route_tests          route ``/process-document`` (per_file, combined, lỗi)
    run_exam_route_tests     route ``/generate-exam`` (plan, bổ sung câu, kẹp count)
    run_concept_tests        route ``/explain-concept`` (4 phần, chuẩn hoá, CORS Vercel)
    run_http_tests           tầng HTTP multipart qua ``TestClient``

Bảo mật cấu hình: key chỉ được đọc từ biến môi trường hoặc file ``.env`` (không
hardcode trong source), thiếu key thì request fail-fast 500, và loader ``.env``
không ghi đè biến môi trường đã có.

Chạy:
    .\\venv\\Scripts\\python.exe test_process_document.py

Phần kiểm thử qua HTTP (``TestClient``) yêu cầu ``httpx``; nếu chưa cài, phần đó
được bỏ qua kèm cảnh báo (các kiểm thử gọi trực tiếp route vẫn chạy bình thường).
Exit code 0 = tất cả pass, 1 = có kiểm thử fail.
"""

from __future__ import annotations

import asyncio
import io
import json
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path
from types import SimpleNamespace

from fastapi import HTTPException
from starlette.datastructures import Headers, UploadFile

import main


# Windows console mặc định không phải UTF-8 -> ép UTF-8 để in được tiếng Việt.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def build_pdf(text: str) -> bytes:
    """Dựng một PDF 1 trang tối giản nhưng hợp lệ để pypdf trích xuất được text."""
    content = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
        b"/Resources << /Font << /F1 5 0 R >> >> >>",
        b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    out = bytearray(b"%PDF-1.4\n")
    offsets: list[int] = []
    for index, body in enumerate(objs, start=1):
        offsets.append(len(out))
        out += f"{index} 0 obj\n".encode() + body + b"\nendobj\n"
    xref_pos = len(out)
    out += f"xref\n0 {len(objs) + 1}\n".encode() + b"0000000000 65535 f \n"
    for offset in offsets:
        out += f"{offset:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode()
    )
    return bytes(out)


def make_upload(filename: str, data: bytes, content_type: str = "application/pdf") -> UploadFile:
    """Tạo ``UploadFile`` trong bộ nhớ giống những gì FastAPI nhận từ multipart."""
    return UploadFile(
        file=io.BytesIO(data),
        filename=filename,
        headers=Headers({"content-type": content_type}),
    )


def build_docx(paragraphs: list[str], table: list[list[str]] | None = None) -> bytes:
    """Dựng file .docx thật (python-docx) gồm đoạn văn + bảng tuỳ chọn."""
    from docx import Document

    document = Document()
    for text in paragraphs:
        document.add_paragraph(text)
    if table:
        grid = document.add_table(rows=len(table), cols=len(table[0]))
        for row_index, row in enumerate(table):
            for column_index, value in enumerate(row):
                grid.cell(row_index, column_index).text = value
    buffer = io.BytesIO()
    document.save(buffer)
    return buffer.getvalue()


def build_xlsx(sheets: dict[str, list[list[object]]]) -> bytes:
    """Dựng file .xlsx thật (openpyxl) với số worksheet / dữ liệu cho trước."""
    from openpyxl import Workbook

    workbook = Workbook()
    workbook.remove(workbook.active)
    for title, rows in sheets.items():
        sheet = workbook.create_sheet(title=title)
        for row in rows:
            sheet.append(row)
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def build_zip(members: dict[str, bytes]) -> bytes:
    """Kho nén hợp lệ nhưng không phải docx/xlsx (dùng để test magic bytes)."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        for name, data in members.items():
            archive.writestr(name, data)
    return buffer.getvalue()


_PLAN_ENTRY_RE = re.compile(r"([a-z_]+)=(\d+)")


def parse_plan(prompt: str) -> list[str]:
    """Đọc dòng "- Số câu theo từng dạng: ..." trong prompt thành plan dạng câu."""
    plan: list[str] = []
    for line in prompt.splitlines():
        if line.startswith("- Số câu theo từng dạng:"):
            for qtype, amount in _PLAN_ENTRY_RE.findall(line):
                plan.extend([qtype] * int(amount))
    return plan


def focus_marker(user_content: str) -> str:
    """MARKER-* của tài liệu chính (per_file) hoặc ghép marker mọi tài liệu (combined)."""
    lines = user_content.splitlines()
    for index, line in enumerate(lines):
        if line.startswith("### TÀI LIỆU CHÍNH"):
            for follow in lines[index + 1 :]:
                if follow.startswith("MARKER-"):
                    return follow.strip()
                if follow.startswith("### "):
                    break
    markers = [line.strip() for line in lines if line.startswith("MARKER-")]
    return "+".join(markers) or "UNKNOWN"


def make_question(qtype: str, index: int, marker: str = "") -> dict:
    """Một câu hỏi hợp lệ theo hợp đồng mới, đúng dạng yêu cầu (dùng cho fake client)."""
    tag = f"{marker} #{index}".strip()
    if qtype == "multiple_choice":
        return {
            "id": index,
            "type": qtype,
            "question": f"Trắc nghiệm {tag}",
            "options": {"A": "Đúng", "B": "Sai", "C": "Sai", "D": "Sai"},
            "answer": "A",
            "explanation": "A đúng",
            "difficulty": "Dễ",
        }
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
    if qtype == "true_false":
        return {
            "id": index,
            "type": qtype,
            "question": f"Nhận định {tag}: Tài sản ngắn hạn luôn có thời hạn trên một năm.",
            "answer": "false",
            "explanation": "Tài sản ngắn hạn có thời hạn dưới một năm",
            "difficulty": "Trung bình",
        }
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


class FakeClient:
    """Client DeepSeek giả: trả về JSON tổng hợp nhiều tài liệu, không gọi mạng.

    ``teaching.summary`` lấy từ marker của tài liệu chính nên test khẳng định được
    kết quả của từng file đến đúng từng file. Câu hỏi được sinh đúng theo phần
    "PHÂN BỔ CÂU HỎI BẮT BUỘC" trong system prompt, nên lần gọi đầu tiên đã đủ số
    câu; ``complete=False`` mô phỏng model trả thiếu câu để kiểm thử nhánh 502.
    """

    def __init__(self, *, complete: bool = True) -> None:
        self.calls: list[dict] = []
        self.complete = complete
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs) -> SimpleNamespace:
        self.calls.append(kwargs)
        marker = focus_marker(kwargs["messages"][-1]["content"])
        plan = parse_plan(kwargs["messages"][0]["content"]) if self.complete else []
        return SimpleNamespace(
            choices=[
                SimpleNamespace(
                    message=SimpleNamespace(
                        content=json.dumps(
                            {
                                "title": f"Tổng hợp {marker}",
                                "teaching": {"summary": marker},
                                "questions": [
                                    make_question(qtype, index, marker)
                                    for index, qtype in enumerate(plan, start=1)
                                ],
                            }
                        )
                    )
                )
            ],
            model="deepseek-chat",
            usage=SimpleNamespace(prompt_tokens=11, completion_tokens=22, total_tokens=33),
        )


_failures: list[str] = []


def ok(cond: bool, label: str) -> None:
    """Ghi nhận một kiểm thử; không dừng ngay để thấy hết lỗi trong một lần chạy."""
    print(("PASS  " if cond else "FAIL  ") + label)
    if not cond:
        _failures.append(label)


async def _missing_key_batch_check() -> None:
    """Thiếu key thì cả request phải 500 ngay, không trả mảng toàn phần tử lỗi."""
    try:
        await main.process_document([make_upload("a.pdf", build_pdf("MARKER-X"))])
    except HTTPException as exc:
        ok(
            exc.status_code == 500 and "DEEPSEEK_API_KEY" in str(exc.detail),
            "thiếu key -> cả batch trả 500 (fail fast)",
        )
    else:
        ok(False, "thiếu key: request phải trả 500")


def run_config_tests() -> None:
    """Bảo đảm API key chỉ đến từ env/``.env``, không bị hardcode trong source."""
    get_client = main.get_deepseek_client
    original_key = os.environ.get("DEEPSEEK_API_KEY")

    try:
        # 1) Không có key -> 500 kèm hướng dẫn cấu hình (không gọi mạng).
        os.environ.pop("DEEPSEEK_API_KEY", None)
        get_client.cache_clear()
        try:
            get_client()
        except HTTPException as exc:
            ok(
                exc.status_code == 500 and "DEEPSEEK_API_KEY" in str(exc.detail),
                "thiếu key -> 500 kèm hướng dẫn cấu hình",
            )
        else:
            ok(False, "thiếu key phải raise HTTPException 500")

        asyncio.run(_missing_key_batch_check())

        # 2) Có key trong env -> dùng đúng key đó và cache lại client.
        os.environ["DEEPSEEK_API_KEY"] = "sk-test-env-123"
        get_client.cache_clear()
        client = get_client()
        ok(getattr(client, "api_key", None) == "sk-test-env-123", "dùng key từ biến môi trường")
        ok(get_client() is client, "client được cache trong process")
    finally:
        get_client.cache_clear()
        if original_key is None:
            os.environ.pop("DEEPSEEK_API_KEY", None)
        else:
            os.environ["DEEPSEEK_API_KEY"] = original_key

    # 3) Loader .env: đọc được, bỏ nháy/tiền tố export, không ghi đè env có sẵn.
    with tempfile.TemporaryDirectory() as tmp_dir:
        env_path = Path(tmp_dir) / ".env"
        env_path.write_text(
            "# comment\n\nFROM_DOTENV=\"gia tri co nhay\"\nexport EXPORTED_DOTENV='single'\n"
            "ALREADY_SET=from-file\n",
            encoding="utf-8",
        )
        for key in ("FROM_DOTENV", "EXPORTED_DOTENV"):
            os.environ.pop(key, None)
        os.environ["ALREADY_SET"] = "from-env"
        try:
            main.load_env_file(env_path)
            ok(os.environ.get("FROM_DOTENV") == "gia tri co nhay", ".env: bỏ dấu nháy kép")
            ok(os.environ.get("EXPORTED_DOTENV") == "single", ".env: hiểu tiền tố export")
            ok(os.environ.get("ALREADY_SET") == "from-env", ".env: không ghi đè env có sẵn")
        finally:
            for key in ("FROM_DOTENV", "EXPORTED_DOTENV", "ALREADY_SET"):
                os.environ.pop(key, None)

    # 4) File .env không tồn tại -> bỏ qua, không lỗi.
    main.load_env_file(Path("khong-ton-tai.env"))
    ok(True, ".env không tồn tại -> bỏ qua an toàn")

    # 5) Không còn API key thật nào nằm trong source code.
    source = Path(main.__file__).read_text(encoding="utf-8")
    ok(re.search(r"sk-[0-9a-zA-Z]{20,}", source) is None, "main.py không còn API key hardcode")


class FakeExamClient:
    """Ghi lại prompt nhận được và trả về một đề giả đúng phân bổ (không gọi mạng)."""

    def __init__(self, content: str | None = None) -> None:
        self.calls: list[dict] = []
        self._content = content
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs) -> SimpleNamespace:
        self.calls.append(kwargs)
        content = self._content
        if content is None:
            plan = parse_plan(kwargs["messages"][0]["content"])
            content = json.dumps(
                {
                    "questions": [
                        make_question(qtype, index)
                        for index, qtype in enumerate(plan, start=1)
                    ]
                }
            )
        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))],
            model="deepseek-chat",
            usage=SimpleNamespace(prompt_tokens=101, completion_tokens=202, total_tokens=303),
        )


class FlakyExamClient:
    """Trả JSON hỏng ở ``broken_calls`` lần đầu rồi mới trả đề hợp lệ (kiểm thử retry)."""

    def __init__(self, *, broken_calls: int) -> None:
        self.calls: list[dict] = []
        self.broken_calls = broken_calls
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs) -> SimpleNamespace:
        self.calls.append(kwargs)
        if len(self.calls) <= self.broken_calls:
            content = "khong phai json"
        else:
            plan = parse_plan(kwargs["messages"][0]["content"])
            content = json.dumps(
                {
                    "questions": [
                        make_question(qtype, index)
                        for index, qtype in enumerate(plan, start=1)
                    ]
                }
            )
        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))],
            model="deepseek-chat",
            usage=SimpleNamespace(prompt_tokens=7, completion_tokens=8, total_tokens=15),
        )


class FakeConceptClient:
    """Client DeepSeek giả cho ``/explain-concept``: trả đủ 4 phần, không gọi mạng."""

    def __init__(self, content: str | None = None) -> None:
        self.calls: list[dict] = []
        self._content = content
        self.chat = SimpleNamespace(completions=SimpleNamespace(create=self._create))

    def _create(self, **kwargs) -> SimpleNamespace:
        self.calls.append(kwargs)
        content = self._content
        if content is None:
            concept = concept_from_prompt(kwargs["messages"][-1]["content"])
            content = json.dumps(
                {
                    "definition": f"Định nghĩa {concept}",
                    "erklaerung": f"Giải thích {concept} kèm ví dụ",
                    "zusammenhang": f"{concept} liên hệ với khái niệm khác",
                    "kritik": f"Ưu nhược điểm của {concept}",
                }
            )
        return SimpleNamespace(
            choices=[SimpleNamespace(message=SimpleNamespace(content=content))],
            model="deepseek-chat",
            usage=SimpleNamespace(prompt_tokens=5, completion_tokens=6, total_tokens=11),
        )


def concept_from_prompt(prompt: str) -> str:
    """Đọc tên khái niệm nằm ngay sau dòng ``===== KHÁI NIỆM CẦN PHÂN TÍCH =====``."""
    lines = prompt.splitlines()
    for index, line in enumerate(lines):
        if line.strip() == "===== KHÁI NIỆM CẦN PHÂN TÍCH =====":
            for follow in lines[index + 1 :]:
                if follow.strip():
                    return follow.strip()
    return "UNKNOWN"


def run_prompt_tests() -> None:
    """Kiểm tra prompt gửi DeepSeek mô tả đúng cấu trúc teaching/questions."""
    user_prompt = main.USER_PROMPT
    system_prompt = main.SYSTEM_PROMPT

    ok("gia sư đại học" in user_prompt, "USER_PROMPT: model đóng vai gia sư đại học")
    ok("teaching" in user_prompt and "questions" in user_prompt, "USER_PROMPT nêu teaching và questions")
    ok(
        "PHÂN BỔ CÂU HỎI BẮT BUỘC" in user_prompt,
        "USER_PROMPT trỏ tới phần phân bổ câu hỏi bắt buộc",
    )
    ok("type" in user_prompt and "table_data" in user_prompt, "USER_PROMPT nêu type và table_data")
    ok("2-3 ví dụ" in user_prompt, "USER_PROMPT yêu cầu 2-3 ví dụ minh hoạ")

    for question_type in (
        "multiple_choice",
        "multiple_response",
        "calculation",
        "fill_in_the_blank",
        "true_false",
        "table",
    ):
        ok(question_type in system_prompt, f"SYSTEM_PROMPT có dạng {question_type}")

    for label in ("trắc nghiệm", "nhiều đáp án", "tính toán", "điền khuyết", "đúng/sai", "kẻ bảng"):
        ok(label in user_prompt, f"USER_PROMPT có dạng {label}")

    ok(
        "ĐÚNG số câu đó" in system_prompt,
        "SYSTEM_PROMPT: số câu do phần phân bổ quy định (không hardcode 10)",
    )

    # --- Phân bổ câu hỏi theo dạng (4 dạng bắt buộc) -------------------------
    plan = main.plan_question_types(10)
    ok(len(plan) == 10, "plan_question_types trả về đúng số câu yêu cầu")
    ok(
        set(plan) == set(main.CORE_QUESTION_TYPES),
        "plan mặc định có đủ 4 dạng bắt buộc",
    )
    counts = main.question_plan_counts(plan)
    ok(sum(counts.values()) == 10, "question_plan_counts cộng đúng tổng số câu")
    ok(
        main.question_plan_counts(main.plan_question_types(8))
        == {"multiple_choice": 2, "calculation": 2, "fill_in_the_blank": 2, "table": 2},
        "plan chia đều câu cho 4 dạng",
    )
    ok(
        main.question_plan_counts(main.plan_question_types(2))
        == {"multiple_choice": 1, "calculation": 1},
        "plan ngắn chỉ lấy các dạng đầu tiên",
    )
    ok(
        main.question_plan_counts(main.plan_question_types(6, "tính toán, table"))
        == {"calculation": 3, "table": 3},
        "plan_question_types lọc theo dạng client chọn (kể cả tiếng Việt)",
    )

    plan_block = main.build_question_plan_block(plan)
    ok(
        "PHÂN BỔ CÂU HỎI BẮT BUỘC" in plan_block and "- Tổng số câu: 10" in plan_block,
        "build_question_plan_block nêu rõ tổng số câu",
    )
    ok(
        "multiple_choice=3" in plan_block,
        "build_question_plan_block liệt kê số câu từng dạng",
    )

    system_with_plan = main.build_system_prompt("vi", plan=plan, cross_document=True)
    ok(
        "PHÂN BỔ CÂU HỎI BẮT BUỘC" in system_with_plan,
        "build_system_prompt chèn phân bổ câu hỏi theo plan",
    )
    ok(
        "YÊU CẦU TỔNG HỢP NHIỀU TÀI LIỆU" in system_with_plan,
        "build_system_prompt thêm luật cross-document",
    )
    ok("TIẾNG VIỆT" in system_with_plan, "build_system_prompt thêm chỉ dẫn ngôn ngữ")

    # --- Prompt cross-document: toàn bộ file của môn nằm trong một lần gọi ---
    documents = [
        main.ExtractedDocument(
            position=1, filename="a.pdf", text="Nội dung A", pages=1, characters=10, truncated=False
        ),
        main.ExtractedDocument(
            position=2, filename="b.pdf", text="Nội dung B", pages=2, characters=20, truncated=False
        ),
    ]
    corpus = main.build_subject_corpus(documents, focus_position=2)
    ok(
        "TÀI LIỆU CHÍNH: b.pdf" in corpus,
        "corpus cross-document đánh dấu tài liệu chính",
    )
    ok(
        "TÀI LIỆU THAM KHẢO CÙNG MÔN: a.pdf" in corpus,
        "corpus cross-document giữ tài liệu tham khảo của môn",
    )
    combined_corpus = main.build_subject_corpus(documents)
    ok(
        "TÀI LIỆU TRONG MÔN: a.pdf" in combined_corpus
        and "TÀI LIỆU TRONG MÔN: b.pdf" in combined_corpus,
        "corpus combined đưa mọi tài liệu vào cùng một khối",
    )
    ok(
        main.build_subject_corpus([], focus_position=1) == "",
        "corpus rỗng khi không có tài liệu",
    )

    document_prompt = main.build_document_user_prompt(
        corpus=corpus,
        plan=plan,
        subject="Kế toán",
        documents=documents,
        focus_filename="b.pdf",
    )
    ok("Môn học: Kế toán" in document_prompt, "build_document_user_prompt nêu tên môn")
    ok(
        "Tài liệu chính của phần này: b.pdf" in document_prompt,
        "build_document_user_prompt nêu tài liệu chính",
    )
    ok(
        "TÀI LIỆU 2: b.pdf (2 trang, 20 ký tự)" in document_prompt,
        "build_document_user_prompt liệt kê mọi file của môn",
    )
    ok("Nội dung A" in document_prompt and "Nội dung B" in document_prompt, "prompt nhúng cả hai tài liệu")

    topup_prompt = main.build_topup_prompt(
        shortfall=2, plan=plan, generated=[], corpus=corpus, subject="Kế toán"
    )
    ok("YÊU CẦU BỔ SUNG CÂU HỎI" in topup_prompt, "build_topup_prompt yêu cầu bổ sung câu thiếu")
    ok("THÊM ĐÚNG 2 câu hỏi MỚI" in topup_prompt, "build_topup_prompt nêu đúng số câu còn thiếu")


    # Khối JSON mẫu trong SYSTEM_PROMPT phải parse được và đúng hợp đồng mới.
    start, end = system_prompt.find("{"), system_prompt.rfind("}")
    schema: dict = json.loads(system_prompt[start : end + 1])

    teaching = schema.get("teaching", {})
    ok(
        {"summary", "key_concepts", "formulas", "examples"} <= set(teaching),
        "SYSTEM_PROMPT: teaching có summary/key_concepts/formulas/examples",
    )
    ok(
        all({"name", "explanation"} <= set(concept) for concept in teaching.get("key_concepts", [])),
        "key_concepts có name + explanation",
    )
    ok(
        all({"formula", "symbols"} <= set(formula) for formula in teaching.get("formulas", [])),
        "formulas có công thức + ký hiệu",
    )
    ok(2 <= len(teaching.get("examples", [])) <= 3, "teaching có 2-3 ví dụ minh hoạ")

    questions = schema.get("questions", [])
    ok(len(questions) == 6, "SYSTEM_PROMPT: JSON mẫu có 6 câu, mỗi dạng 1 câu")
    ok(
        {item.get("type") for item in questions}
        == {
            "multiple_choice",
            "multiple_response",
            "calculation",
            "fill_in_the_blank",
            "true_false",
            "table",
        },
        "SYSTEM_PROMPT: mỗi câu mẫu có trường type đúng dạng",
    )

    # Dạng multiple_response: đáp án là mảng từ 2 chữ cái trở lên, đều có trong options.
    multi_item = next(
        item for item in questions if item.get("type") == "multiple_response"
    )
    multi_answer = multi_item.get("answer", [])
    multi_options = multi_item.get("options", {})
    ok(
        isinstance(multi_answer, list)
        and len(multi_answer) >= 2
        and all(letter in multi_options for letter in multi_answer),
        "multiple_response: answer là mảng >= 2 chữ cái thuộc options",
    )

    # Dạng true_false: đáp án là true/false.
    true_false_item = next(
        item for item in questions if item.get("type") == "true_false"
    )
    ok(
        true_false_item.get("answer") in {"true", "false", True, False},
        "true_false: answer là true/false",
    )

    # Dạng table: table_data gồm headers/rows và ô trống đánh dấu bằng "___".
    table_item = next(item for item in questions if item.get("type") == "table")
    table_data = table_item.get("table_data", {})
    headers = table_data.get("headers", [])
    rows = table_data.get("rows", [])
    ok(bool(headers) and all(isinstance(header, str) for header in headers), "table_data.headers là mảng chuỗi")
    ok(bool(rows) and all(len(row) == len(headers) for row in rows), "table_data.rows cùng độ dài headers")
    ok(any("___" in cell for row in rows for cell in row), "table_data.rows có ô trống ___")
    cell_answers = table_item.get("answer", [])
    ok(
        bool(cell_answers)
        and all({"row", "column", "value"} <= set(answer) for answer in cell_answers),
        "table: answer có row/column/value cho từng ô",
    )

    # --- Prompt tạo đề (/generate-exam) -------------------------------------
    exam_system = main.EXAM_SYSTEM_PROMPT
    for question_type in (
        "multiple_choice",
        "multiple_response",
        "calculation",
        "fill_in_the_blank",
        "true_false",
        "table",
    ):
        ok(question_type in exam_system, f"EXAM_SYSTEM_PROMPT có dạng {question_type}")
    ok("ĐÚNG số lượng câu hỏi" in exam_system, "EXAM_SYSTEM_PROMPT yêu cầu đúng số lượng câu")
    ok("không bịa thêm kiến thức" in exam_system, "EXAM_SYSTEM_PROMPT cấm bịa kiến thức ngoài tài liệu")
    ok("ngẫu nhiên" in exam_system, "EXAM_SYSTEM_PROMPT yêu cầu đáp án ở vị trí ngẫu nhiên")
    ok("TẤT CẢ các chương" in exam_system, "EXAM_SYSTEM_PROMPT yêu cầu phủ mọi chương")

    schema_start = exam_system.find("{")
    schema_end = exam_system.find("\n}\n", schema_start)
    exam_schema = json.loads(exam_system[schema_start : schema_end + 2])
    ok(
        isinstance(exam_schema.get("questions"), list) and exam_schema["questions"],
        "EXAM_SYSTEM_PROMPT: JSON mẫu có mảng questions",
    )

    exam_request = main.ExamRequest(
        context="Chương 1: Lợi nhuận = Doanh thu - Chi phí",
        count=6,
        scope="final",
        chapter_titles=["Chương 1", "Chương 2"],
        existing_questions=["Câu hỏi cũ"],
        seed=123,
        lang="de",
    )
    exam_plan = main.plan_question_types(6, exam_request.types)
    exam_prompt = main.build_exam_prompt(exam_request, main.clamp_exam_count(6), exam_plan)
    ok(
        "PHÂN BỔ CÂU HỎI BẮT BUỘC" in exam_prompt,
        "build_exam_prompt nhúng phân bổ câu hỏi theo dạng",
    )
    ok(
        "PHÂN BỔ CÂU HỎI BẮT BUỘC" in main.build_exam_system_prompt("vi", exam_plan),
        "build_exam_system_prompt nhúng phân bổ câu hỏi theo dạng",
    )
    ok("ĐÚNG 6 câu hỏi" in exam_prompt, "build_exam_prompt nêu đúng số câu")
    ok("Chương 1" in exam_prompt and "Chương 2" in exam_prompt, "build_exam_prompt liệt kê các chương")
    ok("Câu hỏi cũ" in exam_prompt, "build_exam_prompt liệt kê câu hỏi cần tránh lặp")
    ok("Mã biến thể: 123" in exam_prompt, "build_exam_prompt có mã biến thể để tạo đề khác nhau")
    ok("Lợi nhuận" in exam_prompt, "build_exam_prompt nhúng khuôn khổ kiến thức")
    ok("kết thúc học phần" in exam_prompt, "build_exam_prompt mô tả phạm vi final")
    ok(
        "Deutsch" in main.build_exam_system_prompt("de"),
        "build_exam_system_prompt thêm chỉ dẫn tiếng Đức",
    )
    ok(
        "TIẾNG VIỆT" in main.build_exam_system_prompt("vi-VN"),
        "build_exam_system_prompt thêm chỉ dẫn tiếng Việt",
    )


def run_contract_tests() -> None:
    """Kiểm tra các helper chuẩn hoá / phân bổ câu hỏi của hợp đồng mới."""

    # --- teaching: luôn đủ 5 khoá, có fallback summary -----------------------
    teaching = main.normalize_teaching(
        {
            "overview": "Tóm tắt cũ",
            "keyConcepts": [{"name": "Khái niệm", "explanation": "Giải thích"}],
            "formulas": [{"formula": "A = B", "symbols": [{"symbol": "A", "meaning": "..."}]}],
            "examples": ["Ví dụ dạng chuỗi"],
            "crossDocumentLinks": [{"documents": ["a.pdf", "b.pdf"], "relation": "nền tảng"}],
        }
    )
    ok(
        teaching["summary"] == "Tóm tắt cũ"
        and teaching["key_concepts"]
        and teaching["formulas"]
        and teaching["examples"]
        and teaching["cross_document_links"],
        "normalize_teaching đọc được cả khoá camelCase và giữ đủ 5 khoá",
    )
    ok(
        main.normalize_teaching(None, fallback_summary="Dự phòng")["summary"] == "Dự phòng",
        "normalize_teaching dùng fallback summary",
    )
    ok(
        set(main.normalize_teaching({}))
        == {"summary", "key_concepts", "formulas", "examples", "cross_document_links"},
        "normalize_teaching luôn trả đủ 5 khoá",
    )

    # --- questions: chuẩn hoá từng dạng --------------------------------------
    mc = main.normalize_question(
        {
            "type": "MCQ",
            "question": "Câu hỏi",
            "options": ["Đúng", "Sai", "Sai", "Sai"],
            "answer": "a",
        }
    )
    ok(
        mc is not None and mc["type"] == "multiple_choice" and mc["answer"] == "A",
        "MCQ: alias + options mảng",
    )
    ok(
        mc is not None and set(mc["options"]) == {"A", "B", "C", "D"},
        "MCQ: options mảng được đổi thành A/B/C/D",
    )
    ok(
        main.normalize_question({"type": "multiple_choice", "question": "Thiếu options"}) is None,
        "MCQ thiếu options bị loại",
    )
    ok(
        main.normalize_question(
            {"question": "Có bảng", "table_data": {"headers": ["A"], "rows": [["___"]]}}
        )["type"]
        == "table",
        "thiếu type nhưng có table_data -> nhận là table",
    )
    ok(
        main.normalize_question({"question": "Đúng hay sai?", "answer": True})["answer"] == "true",
        "true_false nhận answer boolean",
    )
    ok(
        main.normalize_question({"question": "Điền ___", "answer": "Chi phí"})["answer"] == ["Chi phí"],
        "fill_in_the_blank đổi answer chuỗi thành mảng",
    )
    table = main.normalize_question(
        {
            "type": "table",
            "question": "Điền bảng",
            "table_data": {"headers": ["A", "B"], "rows": [["1", "___"]]},
            "answer": [{"row": 0, "column": 1, "value": "2"}, {"row": "x"}],
        }
    )
    ok(
        table is not None and table["answer"] == [{"row": 0, "column": 1, "value": "2"}],
        "table chỉ giữ ô đáp án hợp lệ",
    )
    ok(
        main.normalize_question({"type": "table", "question": "Bảng rỗng"}) is None,
        "table thiếu table_data bị loại",
    )

    # --- parse_questions / dedupe / align ------------------------------------
    payload = [
        {"type": "multiple_choice", "question": "Câu 1", "options": ["a", "b"], "answer": "A"},
        {"type": "multiple_choice", "question": "Câu  1 ", "options": ["a", "b"], "answer": "A"},
        {"type": "calculation", "question": "Câu 2", "given": ["x"], "answer": "1"},
        {"type": "table", "question": "Câu 3"},
    ]
    parsed = main.parse_questions(payload)
    ok(len(parsed) == 3, "parse_questions nhận mảng thuần và loại câu không hợp lệ")
    ok(len(main.dedupe_questions(parsed)) == 2, "dedupe_questions bỏ câu trùng gần đúng")
    aligned = main.align_questions_to_plan(parsed, ["multiple_choice", "calculation", "table"])
    ok(
        [q["type"] for q in aligned] == ["multiple_choice", "calculation"],
        "align loại câu không khớp plan",
    )
    ok([q["id"] for q in aligned] == [1, 2], "align đánh lại id từ 1")
    ok(
        main.question_type_deficits(aligned, ["multiple_choice", "calculation", "table"])
        == [("table", 1)],
        "question_type_deficits nêu đúng dạng còn thiếu",
    )
    ok(
        main.parse_questions({"questions": "không phải mảng"}) == [],
        "parse_questions chịu được kiểu sai",
    )

    # --- mode / count / lỗi thiếu câu ---------------------------------------
    ok(
        main.parse_document_mode("") == "per_file"
        and main.parse_document_mode("each") == "per_file"
        and main.parse_document_mode("COMBINED") == "combined",
        "parse_document_mode hiểu alias và mặc định per_file",
    )
    try:
        main.parse_document_mode("khong-hop-le")
    except HTTPException as exc:
        ok(exc.status_code == 400, "parse_document_mode từ chối giá trị lạ")
    else:
        ok(False, "parse_document_mode phải raise HTTPException 400")
    ok(
        main.clamp_document_count(0) == main.DOCUMENT_DEFAULT_QUESTIONS
        and main.clamp_document_count(-3) == main.DOCUMENT_DEFAULT_QUESTIONS
        and main.clamp_document_count(999) == main.DOCUMENT_MAX_QUESTIONS,
        "clamp_document_count kẹp đúng biên",
    )
    shortfall = main.QuestionShortfallError(10, 4)
    ok(
        shortfall.requested == 10 and shortfall.produced == 4 and "10" in str(shortfall),
        "QuestionShortfallError giữ số câu yêu cầu / đã tạo",
    )

    # --- tiện ích IO --------------------------------------------------------
    ok(
        main.truncate_text("abcdef", 3) == "abc" and main.truncate_text("abc", 0) == "",
        "truncate_text cắt đúng",
    )
    ok(
        main.merge_usage([])
        == {"prompt_tokens": None, "completion_tokens": None, "total_tokens": None}
        and main.merge_usage(
            [
                {"prompt_tokens": 1, "completion_tokens": 2, "total_tokens": 3},
                {"prompt_tokens": 4, "completion_tokens": 5, "total_tokens": 9},
            ]
        )["total_tokens"]
        == 12,
        "merge_usage cộng token của nhiều lần gọi",
    )
    sources = main.document_sources(
        [
            main.ExtractedDocument(1, "a.pdf", "A", 1, 1, False),
            main.ExtractedDocument(2, "b.pdf", "B", 1, 1, False),
        ],
        focus_position=2,
    )
    ok(
        [source["role"] for source in sources] == ["reference", "focus"],
        "document_sources đánh dấu đúng file tâm điểm",
    )

    # --- JSON robustness: json-repair + max_tokens ---------------------------
    repaired = main.parse_json_response(
        '{"title": "Chương 1", "questions": [{"id": 1, "question": "Tinh'
    )
    ok(
        isinstance(repaired, dict)
        and repaired.get("title") == "Chương 1"
        and repaired.get("questions"),
        "parse_json_response sửa JSON bị cắt cụt bằng json-repair",
    )
    ok(
        main.parse_json_response('Toi la model: {"a": 1, "b": [1, 2')
        == {"a": 1, "b": [1, 2]},
        "parse_json_response cứu JSON có prose + thiếu ngoặc",
    )
    ok(
        main.parse_json_response('```json\n{"a": 1}\n```') == {"a": 1},
        "parse_json_response bỏ markdown fence",
    )
    try:
        main.parse_json_response("khong phai json")
    except json.JSONDecodeError:
        ok(True, "parse_json_response vẫn raise khi không có JSON nào")
    else:
        ok(False, "parse_json_response phải raise JSONDecodeError với prose")
    ok(
        main.DEEPSEEK_MAX_TOKENS >= 8000 and main.EXAM_MAX_TOKENS >= 8000,
        "max_tokens mặc định không thấp hơn 8000",
    )


def run_format_tests() -> None:
    """Đọc nhiều định dạng file: nhận diện đuôi, txt/csv/docx/xlsx và lỗi rõ ràng."""
    # --- nhận diện định dạng ------------------------------------------------
    ok(
        main.SUPPORTED_EXTENSIONS == {".pdf", ".docx", ".xlsx", ".txt", ".csv"},
        "5 định dạng được khai báo hỗ trợ",
    )
    ok(
        main._upload_suffix("BAI-GIANG.PDF") == ".pdf"
        and main._upload_suffix("giao-trinh.docm") == ".docx"
        and main._upload_suffix("diem.xlsm") == ".xlsx"
        and main._upload_suffix("bang-diem.tsv") == ".csv"
        and main._upload_suffix("ghichu.text") == ".txt",
        "_upload_suffix nhận diện đuôi + alias, không phân biệt hoa/thường",
    )
    ok(
        main._upload_suffix("khong-co-duoi", "text/csv") == ".csv"
        and main._upload_suffix("khong-co-duoi", "text/plain; charset=utf-8") == ".txt",
        "_upload_suffix dùng content-type khi tên file không có đuôi",
    )
    for name, content_type, needle in (
        ("bang-tinh.doc", "application/msword", ".docx"),
        ("bang-tinh.xls", "application/vnd.ms-excel", ".xlsx"),
        ("anh.png", "image/png", ".png"),
        ("slide.pptx", "application/octet-stream", ".pptx"),
    ):
        try:
            main._upload_suffix(name, content_type)
        except HTTPException as exc:
            ok(
                exc.status_code == 415 and needle in exc.detail,
                f"{name}: 415 kèm gợi ý {needle}",
            )
        else:
            ok(False, f"{name} phải bị từ chối 415")

    # --- .txt ---------------------------------------------------------------
    text, pages = main.extract_text(
        "bai.txt", "Chương 1 — Kế toán\r\nTài sản".encode("utf-8")
    )
    ok(
        "Kế toán" in text and "\r" not in text and pages == 1,
        "txt: giải mã UTF-8 + chuẩn hoá xuống dòng",
    )
    ok(
        "Hàng tồn kho" in main.extract_text("bai.txt", "Hàng tồn kho".encode("utf-16"))[0],
        "txt: đọc được UTF-16 có BOM",
    )
    ok(
        main.decode_text_bytes("Größe".encode("cp1252")) == "Größe"
        and isinstance(main.decode_text_bytes(bytes(range(256))), str),
        "txt: fallback bảng mã Windows/latin-1 không crash với byte lạ",
    )
    try:
        main.extract_text("bai.txt", b"abc\x00\x01def")
    except main.DocumentFormatError:
        ok(True, "txt: dữ liệu nhị phân -> DocumentFormatError")
    else:
        ok(False, "txt chứa NUL phải bị từ chối")

    # --- .csv ---------------------------------------------------------------
    csv_bytes = (
        "Tài khoản;Số dư;Ghi chú\r\n"
        "131;1.000;Phải thu\r\n"
        '"152";"2,500";"Vật tư, phụ tùng"\r\n'
    ).encode("utf-8")
    csv_text, _ = main.extract_text("so-cai.csv", csv_bytes)
    ok(
        "131 | 1.000 | Phải thu" in csv_text and "152 | 2,500 | Vật tư, phụ tùng" in csv_text,
        "csv: tự dò dấu ';' và render từng dòng thành 'ô | ô'",
    )
    ok(
        main.extract_text("bang.tsv", "Tên\tĐiểm\nAn\t9\n".encode("utf-8"))[0].splitlines()[-1]
        == "An | 9",
        "tsv: tự dò dấu tab",
    )
    ok(
        main.extract_text("a.csv", "Mã,Tên\n01,Chi phí\n".encode("utf-8"))[0].splitlines()[1]
        == "01 | Chi phí",
        "csv: dấu phẩy vẫn được nhận đúng",
    )

    # --- .docx --------------------------------------------------------------
    if main.DocxDocument is None:
        print("SKIP  python-docx chưa cài; bỏ qua kiểm thử .docx")
    else:
        docx_bytes = build_docx(
            ["Chương 2: Hàng tồn kho", "Giá gốc = giá mua + chi phí mua"],
            table=[["Phương pháp", "Công thức"], ["FIFO", "Nhập trước xuất trước"]],
        )
        docx_text, docx_pages = main.extract_text("chuong2.docx", docx_bytes)
        ok(
            "Giá gốc = giá mua" in docx_text
            and "[Bảng 1]" in docx_text
            and "FIFO | Nhập trước xuất trước" in docx_text,
            "docx: đọc được đoạn văn + bảng",
        )
        ok(docx_pages >= 1, "docx: ước lượng được số trang")
        ok(
            main._upload_suffix("chuong2.docx", "") == ".docx",
            "docx: nhận diện theo đuôi file",
        )

    # --- .xlsx --------------------------------------------------------------
    if main.load_xlsx_workbook is None:
        print("SKIP  openpyxl chưa cài; bỏ qua kiểm thử .xlsx")
    else:
        xlsx_bytes = build_xlsx(
            {
                "HangTonKho": [
                    ["Mã", "Tên", "Số lượng"],
                    ["152", "Nguyên vật liệu", 10],
                ],
                "CongThuc": [["Công thức"], ["Giá gốc = giá mua + chi phí mua"]],
            }
        )
        xlsx_text, sheets = main.extract_text("so-lieu.xlsx", xlsx_bytes)
        ok(sheets == 2, "xlsx: số 'trang' = số worksheet")
        ok(
            "### Sheet: HangTonKho" in xlsx_text
            and "152 | Nguyên vật liệu | 10" in xlsx_text
            and "### Sheet: CongThuc" in xlsx_text,
            "xlsx: mỗi sheet thành một khối text dạng bảng",
        )

    # --- file hỏng / sai magic bytes: báo lỗi rõ ràng, không crash -----------
    for name, data in (
        ("hong.docx", b"PK\x03\x04khong-phai-zip"),
        ("anh-gia.xlsx", build_zip({"hello.txt": b"khong phai excel"})),
        ("gia.docx", build_zip({"hello.txt": b"khong phai word"})),
        ("fake.pdf", b"hello world"),
    ):
        try:
            main.extract_text(name, data)
        except main.DocumentFormatError:
            ok(True, f"{name}: nội dung không khớp định dạng -> DocumentFormatError")
        else:
            ok(False, f"{name} phải bị từ chối bằng DocumentFormatError")

    ok(
        "Excel" in main._empty_text_message(".xlsx")
        and "OCR" in main._empty_text_message(".pdf"),
        "_empty_text_message nêu đúng định dạng gặp lỗi",
    )


async def run_route_tests() -> None:
    """Gọi thẳng route ``process_document`` với danh sách ``UploadFile``."""
    fake = FakeClient()
    main.get_deepseek_client = lambda: fake  # type: ignore[assignment]

    pdf_a = build_pdf("MARKER-A1")
    pdf_b = build_pdf("MARKER-B2")

    # 1) Hai PDF hợp lệ -> mảng JSON 2 phần tử, đúng thứ tự file gửi lên.
    results = await main.process_document(
        [make_upload("a.pdf", pdf_a), make_upload("b.pdf", pdf_b)]
    )
    ok(isinstance(results, list) and len(results) == 2, "trả về mảng có 1 phần tử / file")
    ok([r["filename"] for r in results] == ["a.pdf", "b.pdf"], "giữ đúng thứ tự upload")
    ok(all(r["success"] for r in results), "cả hai file đều thành công")
    ok(results[0]["data"]["teaching"]["summary"] == "MARKER-A1", "kết quả file 1 đến từ file 1")
    ok(results[1]["data"]["teaching"]["summary"] == "MARKER-B2", "kết quả file 2 đến từ file 2")
    ok(results[0]["pages"] == 1 and results[0]["usage"]["total_tokens"] == 33, "có metadata từng file")
    ok(len(fake.calls) == 2, "mỗi file gọi DeepSeek đúng 1 lần (không phải bổ sung)")

    # Hợp đồng mới: số câu, phân bổ 4 dạng, cross-document, sources.
    default_plan = main.plan_question_types(main.DOCUMENT_DEFAULT_QUESTIONS)
    first = results[0]
    ok(first["mode"] == "per_file", "mode mặc định là per_file")
    ok(first["requested_count"] == len(default_plan), "requested_count theo số câu mặc định")
    ok(first["count"] == len(default_plan) and first["complete"] is True, "mỗi kết quả đủ số câu")
    ok(
        first["question_plan"] == main.question_plan_counts(default_plan),
        "question_plan khớp phân bổ đã yêu cầu",
    )
    ok(
        {q["type"] for q in first["data"]["questions"]} == set(main.CORE_QUESTION_TYPES),
        "câu hỏi có đủ 4 dạng bắt buộc",
    )
    ok(
        [q["id"] for q in first["data"]["questions"]] == list(range(1, len(default_plan) + 1)),
        "id câu hỏi chạy liên tục từ 1",
    )
    ok(
        {"summary", "key_concepts", "formulas", "examples", "cross_document_links"}
        <= set(first["data"]["teaching"]),
        "teaching có đủ 5 khoá của hợp đồng",
    )
    ok(
        first["documents_used"] == ["a.pdf", "b.pdf"],
        "documents_used liệt kê mọi tài liệu của môn",
    )
    ok(
        [source["filename"] for source in first["data"]["sources"]] == ["a.pdf", "b.pdf"]
        and first["data"]["sources"][0]["role"] == "focus",
        "data.sources ghi rõ file chính / tham khảo",
    )
    ok(
        "TÀI LIỆU THAM KHẢO CÙNG MÔN: b.pdf" in fake.calls[0]["messages"][-1]["content"],
        "prompt file 1 vẫn chứa tài liệu tham khảo của môn",
    )
    ok(
        "Môn học: " not in fake.calls[0]["messages"][-1]["content"],
        "không có subject thì prompt không bịa tên môn",
    )

    # 2) Một file sai định dạng ở giữa không được làm hỏng cả batch.
    bad = make_upload(
        "scan.png", b"\x89PNG\r\n\x1a\nkhong phai tai lieu", content_type="image/png"
    )
    results = await main.process_document(
        [make_upload("ok-1.pdf", pdf_a), bad, make_upload("ok-2.pdf", pdf_b)]
    )
    ok(len(results) == 3, "vẫn trả về đủ 3 phần tử")
    ok(results[0]["success"] and results[2]["success"], "các file hợp lệ vẫn thành công")
    ok(not results[1]["success"], "file lỗi có success=false")
    ok(
        results[1]["status_code"] == 415 and "không được hỗ trợ" in results[1]["error"],
        "file lỗi giữ mã 415 và nêu rõ định dạng không hỗ trợ",
    )
    ok(len(fake.calls) == 4, "file lỗi không gọi DeepSeek")
    ok(
        results[0]["documents_used"] == ["ok-1.pdf", "ok-2.pdf"],
        "file lỗi không nằm trong ngữ cảnh của môn",
    )

    # 3) Content-type PDF nhưng nội dung không phải PDF -> phần tử 400, batch sống.
    results = await main.process_document(
        [make_upload("fake.pdf", b"hello world", content_type="application/pdf")]
    )
    ok(not results[0]["success"] and results[0]["status_code"] == 400, "nội dung không phải PDF -> 400")

    # 3b) Batch hỗn hợp định dạng: PDF + DOCX + XLSX + TXT + CSV đều đọc được.
    if main.DocxDocument is not None and main.load_xlsx_workbook is not None:
        mixed = await main.process_document(
            [
                make_upload("a.pdf", pdf_a),
                make_upload(
                    "b.docx",
                    build_docx(["MARKER-DOCX"], table=[["Cột", "Giá trị"]]),
                    "application/msword",
                ),
                make_upload(
                    "c.xlsx",
                    build_xlsx({"Bang1": [["MARKER-XLSX"]]}),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
                make_upload("d.txt", b"MARKER-TXT", "text/plain"),
                make_upload("e.csv", b"cot1,cot2\nMARKER,CSV", "text/csv"),
            ]
        )
        ok(
            len(mixed) == 5 and all(item["success"] for item in mixed),
            "batch hỗn hợp 5 định dạng đều thành công",
        )
        ok(
            [item["format"] for item in mixed] == ["pdf", "docx", "xlsx", "txt", "csv"],
            "mỗi phần tử báo đúng format",
        )
        ok(
            [source["format"] for source in mixed[1]["data"]["sources"]]
            == ["pdf", "docx", "xlsx", "txt", "csv"],
            "data.sources giữ format của từng tài liệu",
        )
        prompt = fake.calls[-1]["messages"][-1]["content"]
        ok(
            "MARKER-DOCX" in prompt and "MARKER-XLSX" in prompt and "MARKER-TXT" in prompt,
            "prompt chứa nội dung Word/Excel/text đã trích xuất",
        )
    else:
        print("SKIP  thiếu python-docx/openpyxl; bỏ qua batch hỗn hợp định dạng")

    # 3c) Định dạng cũ / định dạng lạ / file text rỗng -> lỗi rõ ràng, không crash.
    legacy = await main.process_document(
        [make_upload("cu.doc", b"\xd0\xcf\x11\xe0", content_type="application/msword")]
    )
    ok(
        legacy[0]["status_code"] == 415 and ".docx" in legacy[0]["error"],
        "file .doc (bản cũ) -> 415 kèm hướng dẫn lưu lại .docx",
    )
    empty = await main.process_document(
        [make_upload("rong.txt", b"   \n\n", content_type="text/plain")]
    )
    ok(empty[0]["status_code"] == 422, "file text rỗng -> 422")

    # 4) Danh sách rỗng -> 400.
    try:
        await main.process_document([])
    except HTTPException as exc:
        ok(exc.status_code == 400, "batch rỗng -> 400")
    else:
        ok(False, "batch rỗng phải raise HTTPException 400")

    # 5) Vượt MAX_BATCH_FILES -> 400.
    original_limit = main.MAX_BATCH_FILES
    main.MAX_BATCH_FILES = 1
    try:
        try:
            await main.process_document([make_upload("a.pdf", pdf_a), make_upload("b.pdf", pdf_b)])
        except HTTPException as exc:
            ok(exc.status_code == 400 and "tối đa" in exc.detail, "vượt giới hạn batch -> 400")
        else:
            ok(False, "vượt giới hạn batch phải raise HTTPException 400")
    finally:
        main.MAX_BATCH_FILES = original_limit

    # 6) mode=combined + subject + count + types: một kết quả cho cả môn.
    results = await main.process_document(
        [make_upload("a.pdf", pdf_a), make_upload("b.pdf", pdf_b)],
        mode="combined",
        subject="Kế toán tài chính",
        count=6,
        types="tính toán, table",
    )
    ok(len(results) == 1, "mode=combined trả về 1 phần tử cho cả môn")
    combined = results[0]
    ok(
        combined["success"] and combined["mode"] == "combined",
        "phần tử combined thành công và giữ mode",
    )
    ok(combined["subject"] == "Kế toán tài chính", "combined giữ tên môn")
    ok(
        combined["requested_count"] == 6 and combined["count"] == 6,
        "combined tôn trọng count của client",
    )
    ok(
        combined["question_plan"] == {"calculation": 3, "table": 3},
        "types lọc đúng các dạng client chọn",
    )
    ok(combined["pages"] == 2, "combined cộng số trang của mọi file")
    ok(
        {q["type"] for q in combined["data"]["questions"]} == {"calculation", "table"},
        "câu hỏi combined chỉ gồm dạng được chọn",
    )
    ok(
        combined["data"]["sources"] and all(
            source["role"] == "reference" for source in combined["data"]["sources"]
        ),
        "combined không có file tâm điểm",
    )
    ok(
        "TÀI LIỆU TRONG MÔN: a.pdf" in fake.calls[-1]["messages"][-1]["content"]
        and "Môn học: Kế toán tài chính" in fake.calls[-1]["messages"][-1]["content"],
        "prompt combined chứa mọi tài liệu và tên môn",
    )
    ok(combined["data"]["title"] == "Tổng hợp MARKER-A1+MARKER-B2", "combined tổng hợp mọi tài liệu")

    # 7) count được kẹp trong DOCUMENT_MIN/MAX_QUESTIONS.
    results = await main.process_document(
        [make_upload("a.pdf", pdf_a)], count=main.DOCUMENT_MAX_QUESTIONS + 5
    )
    ok(
        results[0]["requested_count"] == main.DOCUMENT_MAX_QUESTIONS,
        "count quá lớn được kẹp về DOCUMENT_MAX_QUESTIONS",
    )
    results = await main.process_document([make_upload("a.pdf", pdf_a)], count=0)
    ok(
        results[0]["requested_count"] == main.DOCUMENT_DEFAULT_QUESTIONS,
        "count=0 dùng số câu mặc định",
    )

    # 8) Model trả thiếu câu -> phần tử 502 kèm số câu thiếu (STRICT_QUESTION_COUNT).
    original_get_client = main.get_deepseek_client
    main.get_deepseek_client = lambda: FakeClient(complete=False)  # type: ignore[assignment]
    try:
        results = await main.process_document([make_upload("a.pdf", pdf_a)], count=4)
        element = results[0]
        ok(
            not element["success"] and element["status_code"] == 502,
            "model trả thiếu câu -> phần tử 502",
        )
        ok(
            element.get("requested_count") == 4 and element.get("count") == 0,
            "phần tử lỗi nêu rõ số câu yêu cầu / đã tạo",
        )
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]

    # 9) mode không hợp lệ -> 400, không gọi model.
    calls_before = len(fake.calls)
    try:
        await main.process_document([make_upload("a.pdf", pdf_a)], mode="khong-hop-le")
    except HTTPException as exc:
        ok(exc.status_code == 400 and "mode" in str(exc.detail), "mode không hợp lệ -> 400")
    else:
        ok(False, "mode không hợp lệ phải raise HTTPException 400")
    ok(len(fake.calls) == calls_before, "mode không hợp lệ không gọi DeepSeek")

    # 10) OpenAPI: field ``files`` là mảng binary (gửi lặp lại được) + field mới.
    schema = main.app.openapi()
    body = schema["paths"]["/process-document"]["post"]["requestBody"]
    payload = body["content"]["multipart/form-data"]["schema"]
    ref = payload.get("$ref")
    props = (
        schema["components"]["schemas"][ref.split("/")[-1]]["properties"]
        if ref
        else payload["properties"]
    )
    files_prop = props["files"]
    items = files_prop["items"]
    is_binary = (
        items.get("format") == "binary"
        or items.get("contentMediaType") == "application/octet-stream"
    )
    ok(files_prop.get("type") == "array" and is_binary, "OpenAPI: files là mảng binary")
    ok(
        {"mode", "subject", "count", "types"} <= set(props),
        "OpenAPI: có field mode/subject/count/types",
    )


def run_exam_route_tests() -> None:
    """Gọi thẳng route ``generate_exam`` với client giả (không gọi DeepSeek)."""
    original_get_client = main.get_deepseek_client
    fake = FakeExamClient()
    main.get_deepseek_client = lambda: fake  # type: ignore[assignment]

    try:
        # 1) Đề hợp lệ: id được đánh lại từ 1, giữ đủ metadata.
        request = main.ExamRequest(
            context="Chương 1: Kế toán tài chính",
            count=2,
            scope="chapter",
            chapter_titles=["Chương 1"],
            lang="vi",
            seed=7,
        )
        result = asyncio.run(main.generate_exam(request))
        ok(result["success"] is True and result["count"] == 2, "generate_exam trả về đủ 2 câu")
        ok([q["id"] for q in result["questions"]] == [1, 2], "generate_exam đánh lại id từ 1")
        ok(result["scope"] == "chapter" and result["seed"] == 7, "generate_exam giữ scope/seed")
        ok(result["usage"]["total_tokens"] == 303, "generate_exam trả metadata usage")
        ok(fake.calls and fake.calls[-1]["temperature"] == main.EXAM_TEMPERATURE, "generate_exam dùng nhiệt độ sáng tạo")
        ok(
            fake.calls[-1]["max_tokens"] >= 8000,
            "generate_exam gọi DeepSeek với max_tokens >= 8000",
        )
        ok(
            "TIẾNG VIỆT" in fake.calls[-1]["messages"][0]["content"],
            "generate_exam gửi chỉ dẫn ngôn ngữ trong system prompt",
        )
        ok(
            "ĐÚNG 2 câu hỏi" in fake.calls[-1]["messages"][-1]["content"],
            "generate_exam yêu cầu đúng số câu trong user prompt",
        )
        ok(
            [q["type"] for q in result["questions"]] == ["multiple_choice", "calculation"],
            "generate_exam trả đúng dạng theo phân bổ",
        )
        ok(
            result["question_plan"] == {"multiple_choice": 1, "calculation": 1}
            and result["complete"] is True,
            "generate_exam nêu rõ phân bổ và đã đủ câu",
        )
        ok(
            "PHÂN BỔ CÂU HỎI BẮT BUỘC" in fake.calls[-1]["messages"][0]["content"],
            "generate_exam gửi phân bổ câu hỏi trong system prompt",
        )
        ok(
            len(result["questions"]) == 2,
            "generate_exam bảo đảm đủ số câu (không cần bổ sung)",
        )

        # 2) count ngoài biên được kẹp lại.
        result = asyncio.run(
            main.generate_exam(main.ExamRequest(context="Nội dung", count=999, scope="final"))
        )
        ok(result["requested_count"] == main.EXAM_MAX_QUESTIONS, "count quá lớn được kẹp về MAX")
        ok(
            len(result["questions"]) == main.EXAM_MAX_QUESTIONS,
            "generate_exam sinh đủ số câu đã kẹp",
        )
        ok("ĐÚNG {} câu hỏi".format(main.EXAM_MAX_QUESTIONS) in fake.calls[-1]["messages"][-1]["content"], "prompt dùng số câu đã kẹp")

        result = asyncio.run(
            main.generate_exam(main.ExamRequest(context="Nội dung", count=0))
        )
        ok(result["requested_count"] == main.EXAM_MIN_QUESTIONS, "count 0 được kẹp về MIN")

        # 3) Thiếu context -> 400, không gọi model.
        calls_before = len(fake.calls)
        try:
            asyncio.run(main.generate_exam(main.ExamRequest(context="   ", count=5)))
        except HTTPException as exc:
            ok(exc.status_code == 400, "thiếu context -> 400")
        else:
            ok(False, "thiếu context phải raise HTTPException 400")
        ok(len(fake.calls) == calls_before, "thiếu context không gọi DeepSeek")

        # 4) Model trả JSON hỏng -> 502 (client hiển thị lỗi + cho thử lại).
        main.get_deepseek_client = lambda: FakeExamClient(content="khong phai json")  # type: ignore[assignment]
        try:
            asyncio.run(main.generate_exam(main.ExamRequest(context="Nội dung", count=5)))
        except HTTPException as exc:
            ok(exc.status_code == 502 and "JSON" in str(exc.detail), "JSON hỏng -> 502")
        else:
            ok(False, "JSON hỏng phải raise HTTPException 502")

        # 5) Model trả mảng rỗng / không có questions -> 502.
        main.get_deepseek_client = lambda: FakeExamClient(content=json.dumps({"questions": []}))  # type: ignore[assignment]
        try:
            asyncio.run(main.generate_exam(main.ExamRequest(context="Nội dung", count=5)))
        except HTTPException as exc:
            ok(exc.status_code == 502, "questions rỗng -> 502")
        else:
            ok(False, "questions rỗng phải raise HTTPException 502")

        # 6) JSON hỏng ở lần đầu -> tự gửi lại, lần 2 hợp lệ thì vẫn thành công.
        flaky = FlakyExamClient(broken_calls=1)
        main.get_deepseek_client = lambda: flaky  # type: ignore[assignment]
        result = asyncio.run(
            main.generate_exam(main.ExamRequest(context="Nội dung", count=2))
        )
        ok(
            result["success"] is True and len(result["questions"]) == 2,
            "JSON hỏng lần đầu -> gửi lại và vẫn tạo được đề",
        )
        ok(len(flaky.calls) == 2, "chỉ gửi lại khi JSON thực sự hỏng")
        ok(
            "QUAN TRỌNG" in flaky.calls[-1]["messages"][-1]["content"],
            "lần gửi lại kèm nhắc model trả JSON hoàn chỉnh",
        )
        ok(
            result["usage"]["total_tokens"] == 30,
            "usage của mọi lần gọi đều được cộng dồn",
        )

        # 7) Hỏng cả 3 lần -> 502 sau đúng JSON_PARSE_RETRY_ATTEMPTS lần gửi lại.
        always_broken = FlakyExamClient(broken_calls=99)
        main.get_deepseek_client = lambda: always_broken  # type: ignore[assignment]
        try:
            asyncio.run(main.generate_exam(main.ExamRequest(context="Nội dung", count=2)))
        except HTTPException as exc:
            ok(exc.status_code == 502 and "JSON" in str(exc.detail), "JSON hỏng 3 lần -> 502")
        else:
            ok(False, "JSON hỏng 3 lần phải raise HTTPException 502")
        ok(
            len(always_broken.calls) == main.JSON_PARSE_RETRY_ATTEMPTS + 1,
            "gửi lại đúng số lần cấu hình (1 lần đầu + 2 retry)",
        )
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]


def run_concept_tests() -> None:
    """Kiểm thử ``/explain-concept``: prompt 4 phần, chuẩn hoá, route và CORS Vercel."""
    # --- prompt: đúng 4 khoá + chỉ dẫn ngôn ngữ -------------------------------
    system_prompt = main.build_concept_system_prompt("vi")
    for section in main.CONCEPT_SECTIONS:
        ok(f'"{section}"' in system_prompt, f"CONCEPT_SYSTEM_PROMPT có khoá {section}")
    schema_text = system_prompt[system_prompt.find("{") : system_prompt.rfind("}") + 1]
    try:
        schema = json.loads(schema_text)
    except json.JSONDecodeError as exc:
        ok(False, f"CONCEPT_SYSTEM_PROMPT: khối JSON mẫu parse được ({exc})")
        schema = {}
    else:
        ok(True, "CONCEPT_SYSTEM_PROMPT: khối JSON mẫu parse được")
    ok(
        set(schema) == set(main.CONCEPT_SECTIONS)
        and all(isinstance(value, str) for value in schema.values()),
        "CONCEPT_SYSTEM_PROMPT: JSON mẫu có đúng 4 khoá, giá trị là chuỗi",
    )
    ok("TIẾNG VIỆT" in system_prompt, "build_concept_system_prompt thêm chỉ dẫn tiếng Việt")
    ok(
        "Deutsch" in main.build_concept_system_prompt("de-DE"),
        "build_concept_system_prompt thêm chỉ dẫn tiếng Đức",
    )

    user_prompt = main.build_concept_user_prompt(
        concept="Khấu hao", context="Giá gốc = nguyên giá - khấu hao lũy kế"
    )
    ok(
        "Khấu hao" in user_prompt and "khấu hao lũy kế" in user_prompt,
        "build_concept_user_prompt nhúng khái niệm + context",
    )
    for section in main.CONCEPT_SECTIONS:
        ok(section in user_prompt, f"build_concept_user_prompt nêu phần {section}")
    ok(
        "Không có tài liệu kèm theo" in main.build_concept_user_prompt(concept="X"),
        "build_concept_user_prompt nói rõ khi thiếu context",
    )

    # --- chuẩn hoá JSON: alias, mảng, object, wrapper -------------------------
    flat = main.normalize_concept_explanation(
        {
            "Definition": "Định nghĩa",
            "Erklärung": ["Giải thích 1", "Giải thích 2"],
            "zusammenhang": {"nền tảng": "khái niệm A"},
            "Kritik": "",
        }
    )
    ok(
        set(flat) == set(main.CONCEPT_SECTIONS)
        and all(isinstance(value, str) for value in flat.values()),
        "normalize_concept_explanation luôn trả đủ 4 khoá dạng chuỗi",
    )
    ok(
        flat["definition"] == "Định nghĩa"
        and flat["erklaerung"] == "Giải thích 1\nGiải thích 2"
        and "nền tảng" in flat["zusammenhang"]
        and flat["kritik"] == "",
        "normalize_concept_explanation đọc Erklärung/mảng/object/rỗng",
    )
    nested = main.normalize_concept_explanation(
        {"concept": {"dinh_nghia": "D", "giai_thich": "E"}}
    )
    ok(
        nested["definition"] == "D" and nested["erklaerung"] == "E",
        "normalize_concept_explanation đọc object bọc ngoài + alias tiếng Việt",
    )
    ok(
        main.normalize_concept_explanation("khong phai json")
        == {section: "" for section in main.CONCEPT_SECTIONS},
        "normalize_concept_explanation chịu được payload không phải dict",
    )
    ok(
        main.normalize_concept_explanation({"text": "Chỉ một đoạn văn"})["erklaerung"]
        == "Chỉ một đoạn văn",
        "normalize_concept_explanation dùng đoạn văn làm erklaerung khi thiếu khoá",
    )

    # --- route với client giả -------------------------------------------------
    original_get_client = main.get_deepseek_client
    fake = FakeConceptClient()
    main.get_deepseek_client = lambda: fake  # type: ignore[assignment]
    try:
        result = asyncio.run(
            main.explain_concept(
                main.ConceptRequest(
                    concept="  Khấu   hao  ",
                    context="Tài liệu: khấu hao đường thẳng",
                    lang="vi",
                )
            )
        )
        ok(result["success"] is True, "explain_concept trả success=true")
        ok(
            result["concept"] == "Khấu hao",
            "explain_concept gộp khoảng trắng trong tên khái niệm",
        )
        ok(
            all(isinstance(result[section], str) for section in main.CONCEPT_SECTIONS),
            "explain_concept trả đủ definition/erklaerung/zusammenhang/kritik",
        )
        ok(
            result["definition"].endswith("Khấu hao"),
            "nội dung trả về đúng khái niệm đã hỏi",
        )
        ok(result["usage"]["total_tokens"] == 11, "explain_concept trả metadata usage")
        ok(
            result["context_characters"] == len("Tài liệu: khấu hao đường thẳng"),
            "explain_concept báo số ký tự context",
        )
        ok(result["context_truncated"] is False, "context ngắn không bị đánh dấu cắt")
        ok(
            fake.calls[-1]["temperature"] == main.EXPLAIN_TEMPERATURE,
            "explain_concept dùng EXPLAIN_TEMPERATURE",
        )
        ok(
            fake.calls[-1]["max_tokens"] == main.EXPLAIN_MAX_TOKENS,
            "explain_concept dùng EXPLAIN_MAX_TOKENS",
        )
        ok(
            fake.calls[-1]["response_format"] == {"type": "json_object"},
            "explain_concept yêu cầu DeepSeek trả JSON object",
        )
        ok(
            "TIẾNG VIỆT" in fake.calls[-1]["messages"][0]["content"],
            "explain_concept gửi chỉ dẫn ngôn ngữ trong system prompt",
        )
        prompt = fake.calls[-1]["messages"][-1]["content"]
        ok(
            "Khấu hao" in prompt and "khấu hao đường thẳng" in prompt,
            "user prompt chứa khái niệm và context",
        )
        # thiếu concept -> 400 và không gọi DeepSeek
        calls_before = len(fake.calls)
        try:
            asyncio.run(main.explain_concept(main.ConceptRequest(concept="   ")))
        except HTTPException as exc:
            ok(
                exc.status_code == 400 and "concept" in str(exc.detail),
                "explain_concept: concept rỗng -> 400",
            )
        else:
            ok(False, "concept rỗng phải raise HTTPException 400")
        ok(len(fake.calls) == calls_before, "concept rỗng không gọi DeepSeek")

        # cắt tên khái niệm và context theo giới hạn cấu hình
        original_name_limit = main.MAX_CONCEPT_NAME_CHARS
        main.MAX_CONCEPT_NAME_CHARS = 5
        try:
            result = asyncio.run(
                main.explain_concept(
                    main.ConceptRequest(concept="Khấu hao tài sản cố định")
                )
            )
        finally:
            main.MAX_CONCEPT_NAME_CHARS = original_name_limit
        ok(
            result["concept"] == "Khấu",
            "explain_concept cắt tên khái niệm theo MAX_CONCEPT_NAME_CHARS",
        )

        original_context_limit = main.MAX_CONCEPT_CONTEXT_CHARS
        main.MAX_CONCEPT_CONTEXT_CHARS = 5
        try:
            result = asyncio.run(
                main.explain_concept(
                    main.ConceptRequest(concept="Khấu hao", context="0123456789")
                )
            )
        finally:
            main.MAX_CONCEPT_CONTEXT_CHARS = original_context_limit
        ok(
            result["context_characters"] == 5 and result["context_truncated"] is True,
            "explain_concept cắt context theo MAX_CONCEPT_CONTEXT_CHARS",
        )
        ok(
            "0123456789" not in fake.calls[-1]["messages"][-1]["content"],
            "prompt chỉ chứa phần context trong giới hạn",
        )

        # model trả JSON hỏng -> 502 sau đúng số lần gửi lại
        broken = FakeConceptClient(content="khong phai json")
        main.get_deepseek_client = lambda: broken  # type: ignore[assignment]
        try:
            asyncio.run(main.explain_concept(main.ConceptRequest(concept="X")))
        except HTTPException as exc:
            ok(
                exc.status_code == 502 and "JSON" in str(exc.detail),
                "explain_concept: JSON hỏng -> 502",
            )
        else:
            ok(False, "JSON hỏng phải raise HTTPException 502")
        ok(
            len(broken.calls) == main.JSON_PARSE_RETRY_ATTEMPTS + 1,
            "explain_concept gửi lại đúng JSON_PARSE_RETRY_ATTEMPTS lần",
        )

        # JSON hợp lệ nhưng không có nội dung -> 502
        empty = FakeConceptClient(content=json.dumps({"foo": "bar"}))
        main.get_deepseek_client = lambda: empty  # type: ignore[assignment]
        try:
            asyncio.run(main.explain_concept(main.ConceptRequest(concept="X")))
        except HTTPException as exc:
            ok(
                exc.status_code == 502 and "definition" in str(exc.detail),
                "explain_concept: thiếu definition/erklaerung -> 502",
            )
        else:
            ok(False, "thiếu definition/erklaerung phải raise HTTPException 502")

        # model chỉ trả 1 phần -> vẫn 200, phần còn lại là chuỗi rỗng
        partial = FakeConceptClient(content=json.dumps({"definition": "Chỉ định nghĩa"}))
        main.get_deepseek_client = lambda: partial  # type: ignore[assignment]
        result = asyncio.run(main.explain_concept(main.ConceptRequest(concept="X")))
        ok(
            result["success"] is True
            and result["definition"] == "Chỉ định nghĩa"
            and result["zusammenhang"] == ""
            and result["kritik"] == "",
            "explain_concept: thiếu phần nào thì trả chuỗi rỗng phần đó",
        )
    finally:
        main.get_deepseek_client = original_get_client  # type: ignore[assignment]

    # --- regex CORS cho Vercel ------------------------------------------------
    pattern = main.CORS_ALLOW_ORIGIN_REGEX
    ok(bool(pattern), "CORS_ALLOW_ORIGIN_REGEX mặc định bật cho Vercel")
    for origin in (
        "https://studygap-ai.vercel.app",
        "https://studygap-ai-git-main-team.vercel.app",
        "https://abc-123.vercel.app",
    ):
        ok(
            re.fullmatch(pattern, origin) is not None,
            f"CORS regex chấp nhận {origin}",
        )
    ok(
        re.fullmatch(pattern, "https://vercel.app.evil.example") is None,
        "CORS regex không chấp nhận domain giả mạo",
    )


def run_http_tests() -> None:
    """Gửi multipart thật qua ASGI để chắc chắn field ``files`` lặp lại được."""
    try:
        from fastapi.testclient import TestClient
    except Exception as exc:  # noqa: BLE001 - httpx là dependency tuỳ chọn của test
        print(f"SKIP  TestClient không dùng được ({exc}); bỏ qua các kiểm thử HTTP")
        return

    pdf_a = build_pdf("MARKER-A1")
    pdf_b = build_pdf("MARKER-B2")

    with TestClient(main.app) as client:
        response = client.post(
            "/process-document",
            files=[
                ("files", ("a.pdf", pdf_a, "application/pdf")),
                ("files", ("b.pdf", pdf_b, "application/pdf")),
            ],
        )
        ok(response.status_code == 200, "HTTP 200 cho batch multipart")
        payload = response.json()
        ok(isinstance(payload, list) and len(payload) == 2, "body là mảng JSON 2 phần tử")
        ok([item["filename"] for item in payload] == ["a.pdf", "b.pdf"], "tầng HTTP giữ đúng thứ tự")
        ok(payload[0]["data"]["teaching"]["summary"] == "MARKER-A1", "HTTP: file 1 -> kết quả 1")

        response = client.post(
            "/process-document",
            files=[
                ("files", ("good.pdf", pdf_a, "application/pdf")),
                ("files", ("notes.txt", b"MARKER-TXT", "text/plain")),
                ("files", ("anh.png", b"\x89PNG\r\n\x1a\n", "image/png")),
            ],
        )
        ok(response.status_code == 200, "vẫn HTTP 200 khi có 1 file không hợp lệ")
        payload = response.json()
        ok(
            len(payload) == 3
            and payload[0]["success"]
            and payload[1]["success"]
            and payload[1]["format"] == "txt",
            "HTTP: file .txt được đọc và báo đúng format",
        )
        ok(not payload[2]["success"], "HTTP: file .png thất bại")
        ok(
            payload[2]["status_code"] == 415,
            "HTTP: định dạng không hỗ trợ báo 415 ngay trong mảng",
        )

        response = client.post("/process-document")
        ok(response.status_code == 422, "không gửi file -> FastAPI trả 422")

        # --- /generate-exam qua HTTP -----------------------------------------
        original_get_client = main.get_deepseek_client
        main.get_deepseek_client = lambda: FakeExamClient()  # type: ignore[assignment]
        try:
            response = client.post(
                "/generate-exam",
                json={
                    "context": "Chương 1: Kế toán tài chính",
                    "count": 2,
                    "scope": "final",
                    "lang": "de",
                    "chapter_titles": ["Chương 1", "Chương 2"],
                    "existing_questions": ["Câu cũ"],
                    "seed": 99,
                },
            )
            ok(response.status_code == 200, "HTTP 200 cho /generate-exam")
            payload = response.json()
            ok(
                payload["success"] and len(payload["questions"]) == 2,
                "HTTP: /generate-exam trả về 2 câu",
            )
            ok(
                [item["id"] for item in payload["questions"]] == [1, 2],
                "HTTP: id câu hỏi được đánh lại",
            )

            response = client.post("/generate-exam", json={"context": "", "count": 5})
            ok(response.status_code == 400, "HTTP: thiếu context -> 400")

            response = client.post("/generate-exam", json={})
            ok(response.status_code == 422, "HTTP: thiếu field context -> 422")
        finally:
            main.get_deepseek_client = original_get_client  # type: ignore[assignment]

        # --- /explain-concept qua HTTP ---------------------------------------
        concept_original_get_client = main.get_deepseek_client
        concept_client = FakeConceptClient()
        main.get_deepseek_client = lambda: concept_client  # type: ignore[assignment]
        try:
            vercel_origin = "https://studygap-ai-git-main-team.vercel.app"
            response = client.post(
                "/explain-concept",
                json={
                    "concept": "Khấu hao",
                    "context": "Tài liệu: khấu hao đường thẳng",
                    "lang": "de",
                },
                headers={"Origin": vercel_origin},
            )
            ok(response.status_code == 200, "HTTP 200 cho /explain-concept")
            payload = response.json()
            ok(
                payload["success"] and payload["concept"] == "Khấu hao",
                "HTTP: /explain-concept trả đúng khái niệm",
            )
            ok(
                all(
                    isinstance(payload[section], str)
                    for section in main.CONCEPT_SECTIONS
                ),
                "HTTP: /explain-concept trả đủ 4 phần dạng chuỗi",
            )
            ok(
                payload["usage"]["total_tokens"] == 11,
                "HTTP: /explain-concept trả metadata usage",
            )
            ok(
                "Deutsch" in concept_client.calls[-1]["messages"][0]["content"],
                "HTTP: lang=de được áp dụng",
            )
            allow_origin = response.headers.get("access-control-allow-origin")
            ok(
                allow_origin in {"*", vercel_origin},
                f"HTTP: origin Vercel gọi được API (Allow-Origin={allow_origin})",
            )

            response = client.options(
                "/explain-concept",
                headers={
                    "Origin": vercel_origin,
                    "Access-Control-Request-Method": "POST",
                    "Access-Control-Request-Headers": "content-type",
                },
            )
            ok(response.status_code == 200, "HTTP: preflight CORS cho Vercel -> 200")
            ok(
                response.headers.get("access-control-allow-methods") is not None,
                "HTTP: preflight khai báo method được phép",
            )

            response = client.post("/explain-concept", json={"concept": "   "})
            ok(response.status_code == 400, "HTTP: concept rỗng -> 400")
            response = client.post("/explain-concept", json={"context": "abc"})
            ok(response.status_code == 422, "HTTP: thiếu field concept -> 422")
        finally:
            main.get_deepseek_client = concept_original_get_client  # type: ignore[assignment]

    # --- CORS regex khi đã siết allow_origins (mô phỏng production) -----------
    from fastapi import FastAPI
    from fastapi.middleware.cors import CORSMiddleware

    locked_app = FastAPI()
    locked_app.add_middleware(
        CORSMiddleware,
        allow_origins=["https://studygap-ai.example.com"],
        allow_origin_regex=main.CORS_ALLOW_ORIGIN_REGEX,
        allow_methods=["POST"],
        allow_headers=["content-type"],
    )

    @locked_app.post("/explain-concept")
    async def _locked_concept() -> dict[str, bool]:
        return {"ok": True}

    vercel_origin = "https://studygap-ai-git-main-team.vercel.app"
    with TestClient(locked_app) as locked_client:
        response = locked_client.options(
            "/explain-concept",
            headers={
                "Origin": vercel_origin,
                "Access-Control-Request-Method": "POST",
                "Access-Control-Request-Headers": "content-type",
            },
        )
        ok(response.status_code == 200, "CORS siết origin: preflight Vercel -> 200")
        ok(
            response.headers.get("access-control-allow-origin") == vercel_origin,
            "CORS siết origin: phản chiếu đúng domain Vercel",
        )

        response = locked_client.options(
            "/explain-concept",
            headers={
                "Origin": "https://evil.example",
                "Access-Control-Request-Method": "POST",
            },
        )
        ok(response.status_code == 400, "CORS siết origin: domain lạ bị chặn")


if __name__ == "__main__":
    run_config_tests()
    run_prompt_tests()
    run_contract_tests()
    run_format_tests()
    asyncio.run(run_route_tests())
    run_exam_route_tests()
    run_concept_tests()
    run_http_tests()

    if _failures:
        print(f"\n{len(_failures)} kiểm thử FAIL:")
        for label in _failures:
            print(f"  - {label}")
        sys.exit(1)
    print("\nALL CHECKS PASSED")
