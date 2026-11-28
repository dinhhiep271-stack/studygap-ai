"""StudyGap AI - FastAPI service that turns course documents into DeepSeek generated curricula.

Định dạng đầu vào (tự nhận diện theo đuôi file / content-type):
    .pdf   (pypdf)            .docx  (python-docx)
    .xlsx  (openpyxl)         .txt   (stdlib, nhiều bảng mã)
    .csv   (stdlib csv)       + alias .docm/.xlsm/.tsv/.text
Mọi định dạng được chuyển thành plain text bởi ``extract_text`` trước khi gửi
DeepSeek, nên phần còn lại của pipeline không phụ thuộc định dạng gốc.

Endpoints:
    POST /process-document   (multipart/form-data, field name: ``files`` - nhiều file)
    POST /generate-exam      (application/json) tạo đề luyện thi mới từ khuôn khổ
                             kiến thức của một chương hoặc của cả môn học
    POST /explain-concept    (application/json) phân tích MỘT khái niệm thành 4 phần
                             definition / erklaerung / zusammenhang / kritik, dựa trên
                             ``context`` là nội dung tài liệu liên quan; mỗi phần là
                             object song ngữ ``{"de": ..., "vi": ...}``

Song ngữ Đức - Việt (bilingual):
    ``teaching.summary_de`` / ``summary_vi`` và ``teaching.concepts_de`` /
    ``concepts_vi`` luôn có đủ hai bản, nên client chỉ cần đọc đúng một lần là có
    cả tiếng Đức và tiếng Việt. Hai khoá ``summary`` / ``key_concepts`` (không hậu
    tố) vẫn được trả về như bản tương ứng với ``lang`` để client cũ không phải sửa
    ngay. ``/explain-concept`` trả mỗi phần trong 4 phần dạng ``{"de", "vi"}``.

Independent per-file analysis (yêu cầu 1):
    Mỗi file trong **một request** được xử lý ĐỘC LẬP: trích xuất text của riêng
    file đó rồi gửi DUY NHẤT nội dung ấy cho DeepSeek. Mỗi lần gọi model đều nhận
    ``===== NỘI DUNG TÀI LIỆU DUY NHẤT =====`` của đúng một file, kèm chỉ dẫn
    ``Chỉ phân tích duy nhất tài liệu này. Không gộp nội dung với bất kỳ tài liệu
    nào khác. Tạo tóm tắt và khái niệm riêng cho chương này.`` Vì vậy ``teaching``
    và câu hỏi của mỗi phần tử chỉ phản ánh file tương ứng, không trộn nội dung
    của các file khác.

    - ``mode=per_file`` (mặc định): một phần tử / file, mỗi phần tử phân tích độc
      lập đúng file của nó.
    - ``mode=combined`` (tuỳ chọn, tương thích ngược): MỘT phần tử tổng hợp cho cả
      môn học, chỉ dùng khi client chủ động yêu cầu.

Exercise generation (yêu cầu 2-4):
    1. ``plan_question_types`` chia chính xác số câu người dùng yêu cầu cho 4 dạng
       bắt buộc: trắc nghiệm, tính toán, điền khuyết, kẻ bảng.
    2. Mỗi câu hỏi có ``id`` đánh số liên tục từ 1.
    3. ``complete_questions`` đếm số câu trả về; nếu thiếu thì gọi DeepSeek thêm
       (tối đa ``QUESTION_TOPUP_ATTEMPTS`` lần) và chỉ trả về khi đủ số câu, hoặc
       báo 502 khi ``STRICT_QUESTION_COUNT`` bật mà model vẫn không đủ.
    4. JSON hỏng/bị cắt cụt được ``parse_json_response`` sửa bằng ``json-repair``;
       nếu vẫn không parse được, ``request_json_payload`` gửi lại yêu cầu cho
       DeepSeek tối đa ``JSON_PARSE_RETRY_ATTEMPTS`` lần (mặc định 2) trước khi báo lỗi.

Run locally (PowerShell):
    $env:DEEPSEEK_API_KEY = "sk-..."
    .\\venv\\Scripts\\python.exe -m uvicorn main:app --reload --port 8000

Hoặc tạo file .env (đã có trong .gitignore) từ .env.example:
    DEEPSEEK_API_KEY=sk-...

Interactive docs: http://127.0.0.1:8000/docs
"""

from __future__ import annotations

import csv
import datetime as dt
import io
import json
import logging
import math
import os
import threading
import uuid
import zipfile
from collections import Counter
from dataclasses import dataclass
from dotenv import load_dotenv
load_dotenv()
import re
from functools import lru_cache
from pathlib import Path
from typing import Annotated, Any, Callable, Iterable, Sequence

from fastapi import (
    BackgroundTasks,
    FastAPI,
    File,
    Form,
    HTTPException,
    Response,
    UploadFile,
    status,
)
from fastapi.middleware.cors import CORSMiddleware
from openai import APIConnectionError, APIStatusError, APITimeoutError, OpenAI, OpenAIError
from pydantic import BaseModel, Field
from pypdf import PdfReader
from pypdf.errors import PdfReadError

# json-repair tự đóng ngoặc/chuỗi còn thiếu nên cứu được JSON bị model cắt cụt.
# Vẫn import "mềm" để server khởi động được nếu môi trường chưa cài thư viện.
try:  # json-repair -> parse_json_response
    from json_repair import repair_json as _json_repair
except ImportError:  # pragma: no cover - phụ thuộc vào môi trường cài đặt
    _json_repair = None  # type: ignore[assignment]

# Đọc Word/Excel là tuỳ chọn: nếu thư viện chưa cài, server vẫn chạy và chỉ báo
# lỗi rõ ràng (kèm lệnh pip install) khi có người gửi đúng định dạng đó.
try:  # python-docx -> file .docx / .docm
    from docx import Document as DocxDocument
except ImportError:  # pragma: no cover - phụ thuộc vào môi trường cài đặt
    DocxDocument = None  # type: ignore[assignment]

try:  # openpyxl -> file .xlsx / .xlsm
    from openpyxl import load_workbook as load_xlsx_workbook
except ImportError:  # pragma: no cover - phụ thuộc vào môi trường cài đặt
    load_xlsx_workbook = None  # type: ignore[assignment]

# --------------------------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
)
logger = logging.getLogger("studygap.document")

# --------------------------------------------------------------------------------------
# .env (optional, stdlib only)
# --------------------------------------------------------------------------------------
ENV_FILE = Path(__file__).with_name(".env")


def load_env_file(path: Path = ENV_FILE) -> None:
    """Nạp các cặp ``KEY=VALUE`` từ file ``.env`` nếu file tồn tại.

    Không cần thêm dependency nào. Biến đã có trong môi trường (ví dụ đặt bằng
    ``$env:DEEPSEEK_API_KEY``) luôn được ưu tiên; file chỉ bổ sung biến còn thiếu.
    Hỗ trợ dòng trống, comment ``#``, tiền tố ``export`` và giá trị trong nháy.
    File ``.env`` đã nằm trong ``.gitignore`` nên không bị commit.
    """
    if not path.is_file():
        return
    try:
        # utf-8-sig để chịu được BOM do Notepad/PowerShell tạo ra.
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    except OSError:
        logger.warning("Không đọc được file %s", path, exc_info=True)
        return

    loaded: list[str] = []
    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("export "):
            line = line[len("export ") :].lstrip()
        key, separator, value = line.partition("=")
        key = key.strip()
        if not separator or not key or key in os.environ:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        os.environ[key] = value
        loaded.append(key)

    if loaded:
        # Chỉ log tên biến, không bao giờ log giá trị (có thể là API key).
        logger.info("Đã nạp biến môi trường từ %s: %s", path, ", ".join(loaded))


load_env_file()


# --------------------------------------------------------------------------------------
# Configuration (all values can be overridden with environment variables)
# --------------------------------------------------------------------------------------
DEEPSEEK_BASE_URL = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
DEEPSEEK_MODEL = os.getenv("DEEPSEEK_MODEL", "deepseek-chat")
DEEPSEEK_TIMEOUT_SECONDS = float(os.getenv("DEEPSEEK_TIMEOUT_SECONDS", "300"))
# 8000 tokens is the safe floor: curricula / exams with many questions used to be
# truncated mid-JSON at smaller limits. Env can still RAISE it, never lower it.
MIN_COMPLETION_TOKENS = 8000
DEEPSEEK_MAX_TOKENS = max(
    MIN_COMPLETION_TOKENS, int(os.getenv("DEEPSEEK_MAX_TOKENS", "8000"))
)
DEEPSEEK_TEMPERATURE = float(os.getenv("DEEPSEEK_TEMPERATURE", "0.3"))
# Exam generation is deliberately more creative than document extraction: the
# same chapter must produce a different paper on every click.
EXAM_TEMPERATURE = float(os.getenv("EXAM_TEMPERATURE", "0.9"))
EXAM_MAX_TOKENS = max(MIN_COMPLETION_TOKENS, int(os.getenv("EXAM_MAX_TOKENS", "8000")))

# JSON robustness: a response that is still broken after json-repair is requested
# again, up to this many extra times (2 => at most 3 DeepSeek calls per payload).
JSON_PARSE_RETRY_ATTEMPTS = max(0, int(os.getenv("JSON_PARSE_RETRY_ATTEMPTS", "2")))

# Guards so a huge upload cannot blow up the prompt / the request payload.
# MAX_PDF_BYTES là tên cũ, vẫn đọc được để không phá vỡ cấu hình đã có.
MAX_UPLOAD_BYTES = int(
    os.getenv("MAX_UPLOAD_BYTES") or os.getenv("MAX_PDF_BYTES") or str(20 * 1024 * 1024)
)  # 20 MB
MAX_PDF_BYTES = MAX_UPLOAD_BYTES  # alias tương thích ngược
MAX_TEXT_CHARS = int(os.getenv("MAX_TEXT_CHARS", "120000"))
# Upper bound for a single batch request: each file costs one DeepSeek call.
MAX_BATCH_FILES = int(os.getenv("MAX_BATCH_FILES", "10"))
# Chặn "zip bomb": file .docx/.xlsx nén nhỏ nhưng giải nén ra rất lớn.
MAX_ZIP_EXPANSION_RATIO = int(os.getenv("MAX_ZIP_EXPANSION_RATIO", "50"))
# DOCX/TXT/CSV không có khái niệm trang -> ước lượng số trang theo số ký tự.
CHARS_PER_PAGE_ESTIMATE = int(os.getenv("CHARS_PER_PAGE_ESTIMATE", "3000"))

# Document analysis budget: "per_file" (default) sends ONLY the current file to
# the model (max MAX_TEXT_CHARS characters), so each result is independent.
# "combined" (opt-in) shares MAX_SUBJECT_CONTEXT_CHARS between all files;
# MAX_CROSS_DOCUMENT_CHARS is kept for that legacy mode.
MAX_SUBJECT_CONTEXT_CHARS = int(os.getenv("MAX_SUBJECT_CONTEXT_CHARS", "100000"))
MAX_CROSS_DOCUMENT_CHARS = int(os.getenv("MAX_CROSS_DOCUMENT_CHARS", "30000"))
MIN_DOCUMENT_EXCERPT_CHARS = int(os.getenv("MIN_DOCUMENT_EXCERPT_CHARS", "1200"))

# /process-document question plan. `count=0` (the default form value) means
# "use DOCUMENT_DEFAULT_QUESTIONS"; the value is clamped to the MIN/MAX bounds.
DOCUMENT_MIN_QUESTIONS = int(os.getenv("DOCUMENT_MIN_QUESTIONS", "1"))
DOCUMENT_DEFAULT_QUESTIONS = int(os.getenv("DOCUMENT_DEFAULT_QUESTIONS", "10"))
DOCUMENT_MAX_QUESTIONS = int(os.getenv("DOCUMENT_MAX_QUESTIONS", "30"))

# Quality control (requirement 4): a short answer is topped up with extra model
# calls. STRICT_QUESTION_COUNT=True guarantees the client never receives fewer
# questions than requested (a 502 is raised instead).
QUESTION_TOPUP_ATTEMPTS = int(os.getenv("QUESTION_TOPUP_ATTEMPTS", "3"))
STRICT_QUESTION_COUNT = os.getenv("STRICT_QUESTION_COUNT", "true").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}

# /generate-exam limits. The Next.js API already clamps the values, these are
# the last line of defence for other clients of the service.
EXAM_MIN_QUESTIONS = int(os.getenv("EXAM_MIN_QUESTIONS", "1"))
EXAM_MAX_QUESTIONS = int(os.getenv("EXAM_MAX_QUESTIONS", "30"))
MAX_EXAM_CONTEXT_CHARS = int(os.getenv("MAX_EXAM_CONTEXT_CHARS", "16000"))
MAX_EXAM_EXISTING_QUESTIONS = 15

# /explain-concept: one concept = one short call. A lower temperature keeps the
# explanation faithful to the documents; the context is the excerpt the frontend
# sends along, so the answer stays grounded in the subject material.
EXPLAIN_TEMPERATURE = float(os.getenv("EXPLAIN_TEMPERATURE", "0.4"))
# Four prose sections do not need the 8000-token floor of a full exam, but the
# env may still raise the limit. Song ngữ nhân đôi lượng chữ (4 phần × de + vi)
# nên mặc định 6000 để câu trả lời dài không bị cắt cụt (floor 2000).
EXPLAIN_MAX_TOKENS = max(2000, int(os.getenv("EXPLAIN_MAX_TOKENS", "6000")))
MAX_CONCEPT_CONTEXT_CHARS = int(os.getenv("MAX_CONCEPT_CONTEXT_CHARS", "30000"))
MAX_CONCEPT_NAME_CHARS = int(os.getenv("MAX_CONCEPT_NAME_CHARS", "200"))

# /process-book: chia một quyển sách PDF thành từng chương dựa trên mục lục
# (Inhaltsverzeichnis) rồi xử lý từng chương bằng đúng pipeline của
# /process-document. Mục lục chỉ được dò ở phần đầu sách để không nhầm với chữ
# "Inhalt" trong nội dung; mỗi chương tốn ít nhất một lần gọi DeepSeek nên số
# chương và số trang dò mục lục đều có giới hạn.
BOOK_TOC_SEARCH_PAGES = max(1, int(os.getenv("BOOK_TOC_SEARCH_PAGES", "40")))
BOOK_TOC_MAX_PAGES = max(1, int(os.getenv("BOOK_TOC_MAX_PAGES", "12")))
# Số trang in trên mục lục có thể lớn hơn số trang PDF (trang bìa, phụ lục...).
BOOK_TOC_PAGE_TOLERANCE = max(0, int(os.getenv("BOOK_TOC_PAGE_TOLERANCE", "50")))
MAX_BOOK_CHAPTERS = max(1, int(os.getenv("MAX_BOOK_CHAPTERS", "40")))
# Không tìm được mục lục/tiêu đề chương: chia đều sách theo số trang này.
BOOK_FALLBACK_CHAPTER_PAGES = max(1, int(os.getenv("BOOK_FALLBACK_CHAPTER_PAGES", "20")))
# Job chạy ngầm giữ kết quả trong bộ nhớ process (Render): TTL + số job tối đa
# để một quyển sách dài không làm phình RAM.
BOOK_JOB_TTL_SECONDS = max(60, int(os.getenv("BOOK_JOB_TTL_SECONDS", "86400")))
MAX_BOOK_JOBS = max(1, int(os.getenv("MAX_BOOK_JOBS", "25")))
# Mặc định xử lý ngầm để không bị timeout trên Render; client gửi
# ``background=false`` sẽ nhận ngay mảng kết quả trong một response.
BOOK_BACKGROUND_BY_DEFAULT = os.getenv("BOOK_BACKGROUND_BY_DEFAULT", "true").strip().lower() not in {
    "0",
    "false",
    "no",
    "off",
}


def _http_status(new_name: str, legacy_name: str) -> int:
    """Starlette renamed a few codes (RFC 9110); stay compatible with both."""
    value = getattr(status, new_name, None)
    return value if value is not None else getattr(status, legacy_name)


HTTP_413_CONTENT_TOO_LARGE = _http_status(
    "HTTP_413_CONTENT_TOO_LARGE", "HTTP_413_REQUEST_ENTITY_TOO_LARGE"
)
HTTP_422_UNPROCESSABLE_CONTENT = _http_status(
    "HTTP_422_UNPROCESSABLE_CONTENT", "HTTP_422_UNPROCESSABLE_ENTITY"
)

# Allow All Origins: mặc định "*" để frontend (ví dụ trang web deploy trên Vercel,
# hoặc bất kỳ domain nào) gọi API mà không bị trình duyệt chặn CORS.
# Chỉ khi thật sự muốn siết lại mới đặt CORS_ALLOW_ORIGINS="https://a.vercel.app,https://b.com"
# (danh sách phân tách bằng dấu phẩy); biến trống hoặc không đặt => "*".
_DEFAULT_ALLOWED_ORIGINS = ["*"]
ALLOWED_ORIGINS = [
    origin.strip()
    for origin in os.getenv(
        "CORS_ALLOW_ORIGINS", ",".join(_DEFAULT_ALLOWED_ORIGINS)
    ).split(",")
    if origin.strip()
] or list(_DEFAULT_ALLOWED_ORIGINS)
# True khi mọi origin đều được phép (dùng cho log khởi động).
CORS_ALLOW_ALL_ORIGINS = "*" in ALLOWED_ORIGINS

# Domain Vercel (production lẫn preview dạng https://<project>-<hash>-<team>.vercel.app)
# luôn được phép, kể cả khi CORS_ALLOW_ORIGINS đã bị siết lại. Starlette KHÔNG hiểu
# wildcard trong allow_origins nên phải dùng allow_origin_regex (fullmatch trên Origin).
# Đặt CORS_ALLOW_ORIGIN_REGEX="" để tắt hoàn toàn regex này.
_DEFAULT_ALLOWED_ORIGIN_REGEX = r"https://([a-z0-9-]+\.)*vercel\.app"
CORS_ALLOW_ORIGIN_REGEX = os.getenv(
    "CORS_ALLOW_ORIGIN_REGEX", _DEFAULT_ALLOWED_ORIGIN_REGEX
).strip()

# --------------------------------------------------------------------------------------
# Question types and per-request plan (requirements 2-4)
# --------------------------------------------------------------------------------------
# Four exercise formats are mandatory in every generated exam. The two extra
# formats stay supported so already stored curricula keep rendering correctly.
CORE_QUESTION_TYPES: tuple[str, ...] = (
    "multiple_choice",
    "calculation",
    "fill_in_the_blank",
    "table",
)
SUPPORTED_QUESTION_TYPES: tuple[str, ...] = (
    "multiple_choice",
    "multiple_response",
    "calculation",
    "fill_in_the_blank",
    "true_false",
    "table",
)
QUESTION_TYPE_LABELS = {
    "multiple_choice": "trắc nghiệm một đáp án",
    "multiple_response": "trắc nghiệm nhiều đáp án",
    "calculation": "tính toán",
    "fill_in_the_blank": "điền khuyết",
    "true_false": "đúng/sai",
    "table": "kẻ bảng",
}

# The model (and older clients) do not always use the canonical enum, so every
# spelling we have seen is mapped back to the canonical value.
_QUESTION_TYPE_ALIASES: dict[str, str] = {
    "multiple choice": "multiple_choice",
    "multiplechoice": "multiple_choice",
    "multiple choices": "multiple_choice",
    "single choice": "multiple_choice",
    "single answer": "multiple_choice",
    "one answer": "multiple_choice",
    "mcq": "multiple_choice",
    "mc": "multiple_choice",
    "multiple response": "multiple_response",
    "multiple responses": "multiple_response",
    "multiple answers": "multiple_response",
    "multi choice": "multiple_response",
    "multi select": "multiple_response",
    "calculation": "calculation",
    "calc": "calculation",
    "computation": "calculation",
    "problem": "calculation",
    "exercise": "calculation",
    "fill in the blank": "fill_in_the_blank",
    "fill in the blanks": "fill_in_the_blank",
    "fill in blank": "fill_in_the_blank",
    "fill blank": "fill_in_the_blank",
    "fill blanks": "fill_in_the_blank",
    "blank": "fill_in_the_blank",
    "cloze": "fill_in_the_blank",
    "true false": "true_false",
    "true or false": "true_false",
    "truefalse": "true_false",
    "boolean": "true_false",
    "table": "table",
    "table completion": "table",
    "fill table": "table",
    "table fill": "table",
    "grid": "table",
}
_VIETNAMESE_TYPE_ALIASES: dict[str, str] = {
    "trắc nghiệm": "multiple_choice",
    "trắc nghiệm một đáp án": "multiple_choice",
    "trắc nghiệm 1 đáp án": "multiple_choice",
    "trắc nghiệm nhiều đáp án": "multiple_response",
    "nhiều đáp án": "multiple_response",
    "tính toán": "calculation",
    "bài toán": "calculation",
    "điền khuyết": "fill_in_the_blank",
    "điền vào chỗ trống": "fill_in_the_blank",
    "đúng sai": "true_false",
    "đúng/sai": "true_false",
    "kẻ bảng": "table",
    "điền bảng": "table",
    "bảng": "table",
}


def normalize_question_type(value: Any) -> str | None:
    """Map an arbitrary label ("MCQ", "tính toán", ...) to a canonical type."""
    if value is None:
        return None
    text = " ".join(str(value).replace("_", " ").replace("-", " ").split()).lower()
    if not text:
        return None
    return _QUESTION_TYPE_ALIASES.get(text) or _VIETNAMESE_TYPE_ALIASES.get(text)


def parse_question_types(raw: str | Iterable[Any] | None) -> list[str]:
    """Parse ``"multiple_choice, tính toán"`` or ``["mcq", ...]`` into types."""
    if raw is None:
        return []
    if isinstance(raw, str):
        tokens: list[Any] = raw.replace(";", ",").replace("|", ",").split(",")
    else:
        tokens = list(raw)
    types: list[str] = []
    for token in tokens:
        canonical = normalize_question_type(token)
        if canonical and canonical not in types:
            types.append(canonical)
    return types


def plan_question_types(
    count: int, types: str | Iterable[Any] | None = None
) -> list[str]:
    """Return exactly ``count`` question types, one per question.

    Types are distributed round-robin so the four mandatory formats appear in a
    predictable order (MC, calculation, fill-in-the-blank, table). Asking for
    fewer questions than formats keeps only the first formats, because a single
    question cannot cover two formats.
    """
    total = max(0, int(count))
    if total == 0:
        return []
    selected = [t for t in parse_question_types(types) if t in SUPPORTED_QUESTION_TYPES]
    if not selected:
        selected = list(CORE_QUESTION_TYPES)
    return [selected[index % len(selected)] for index in range(total)]


def question_plan_counts(plan: Sequence[str]) -> dict[str, int]:
    """Return ``{type: count}`` in canonical type order, empty entries dropped."""
    counts = Counter(plan)
    return {t: counts[t] for t in SUPPORTED_QUESTION_TYPES if counts.get(t)}


def build_question_plan_block(plan: Sequence[str]) -> str:
    """Render the machine-checkable "exactly N questions" instruction."""
    total = len(plan)
    counts = question_plan_counts(plan)
    lines = [
        "===== PHÂN BỔ CÂU HỎI BẮT BUỘC =====",
        f"- Tổng số câu: {total}",
        "- Số câu theo từng dạng: "
        + "; ".join(f"{qtype}={amount}" for qtype, amount in counts.items()),
        f'- id chạy liên tục từ 1 đến {total}; mỗi câu có trường "type" đúng dạng đã phân bổ.',
    ]
    if total >= len(CORE_QUESTION_TYPES):
        lines.append(
            "- Bắt buộc có đủ 4 dạng: "
            + ", ".join(
                f"{qtype} ({QUESTION_TYPE_LABELS[qtype]})"
                for qtype in CORE_QUESTION_TYPES
            )
            + "."
        )
    return "\n".join(lines)


# Prompt gửi cho DeepSeek: trả về 2 phần teaching (giảng dạy tổng hợp nhiều tài liệu)
# và questions (bài tập thuộc 4 dạng bắt buộc multiple_choice / calculation /
# fill_in_the_blank / table, cùng 2 dạng mở rộng multiple_response / true_false).
USER_PROMPT = """Bạn là một gia sư đại học tận tâm phụ trách tổng hợp TẤT CẢ tài liệu của cùng một môn học. Hãy đọc kỹ toàn bộ tài liệu được cung cấp (có thể gồm nhiều file) rồi trả về JSON với cấu trúc sau:
1. teaching: Phần giảng dạy tổng hợp SONG NGỮ ĐỨC - VIỆT, bao gồm:
   - summary_de và summary_vi: Tóm tắt nội dung chính của cả môn/chương, nêu rõ tài liệu nào bổ sung ý nào; summary_de viết bằng TIẾNG ĐỨC, summary_vi viết bằng TIẾNG VIỆT, hai bản tương đương về ý.
   - concepts_de và concepts_vi: Danh sách khái niệm quan trọng (kèm giải thích ngắn gọn và tên file chứa khái niệm); concepts_de viết bằng TIẾNG ĐỨC, concepts_vi viết bằng TIẾNG VIỆT, cùng số lượng và cùng thứ tự khái niệm.
   - formulas: Danh sách công thức (nếu có), kèm giải thích ý nghĩa từng ký hiệu.
   - examples: 2-3 ví dụ minh họa cụ thể, có lời giải chi tiết.
   - cross_document_links: Liên kết kiến thức giữa các tài liệu (tài liệu nào là nền tảng cho tài liệu nào).
2. questions: Phần bài tập, gồm đủ 4 dạng: trắc nghiệm một đáp án, tính toán, điền khuyết và kẻ bảng (có thể thêm trắc nghiệm nhiều đáp án và đúng/sai nếu phù hợp). Mỗi câu hỏi có trường type để phân biệt và trường id đánh số từ 1. Với dạng table, trả về table_data gồm headers và rows (dùng ___ cho ô trống).

Số lượng câu hỏi và số câu của từng dạng được quy định chính xác ở phần "PHÂN BỔ CÂU HỎI BẮT BUỘC" bên dưới — phải tuân thủ đúng, không tự ý giảm số câu."""

# Prompt dùng cho chế độ mặc định ``per_file``: mỗi lần gọi DeepSeek chỉ nhận nội
# dung của MỘT file, nên phần giảng dạy phải là của riêng chương đó. Nội dung
# ``SYSTEM_PROMPT`` / cấu trúc JSON vẫn giữ nguyên để client không phải sửa.
SINGLE_DOCUMENT_USER_PROMPT = """Bạn là một gia sư đại học tận tâm. Hãy CHỈ phân tích DUY NHẤT tài liệu được cung cấp trong yêu cầu này (một chương) rồi trả về JSON với cấu trúc sau:
1. teaching: Phần giảng dạy của RIÊNG chương này, SONG NGỮ ĐỨC - VIỆT, bao gồm:
   - summary_de và summary_vi: Tóm tắt nội dung chính của DUY NHẤT chương này; summary_de viết bằng TIẾNG ĐỨC, summary_vi viết bằng TIẾNG VIỆT, hai bản tương đương về ý.
   - concepts_de và concepts_vi: Danh sách khái niệm quan trọng CỦA CHƯƠNG NÀY (kèm giải thích ngắn gọn và tên file chứa khái niệm); concepts_de viết bằng TIẾNG ĐỨC, concepts_vi viết bằng TIẾNG VIỆT, cùng số lượng và cùng thứ tự khái niệm.
   - formulas: Danh sách công thức có trong chương này (nếu có), kèm giải thích ý nghĩa từng ký hiệu.
   - examples: 2-3 ví dụ minh họa cụ thể của chương này, có lời giải chi tiết.
   - cross_document_links: Luôn là mảng rỗng [] vì kết quả này chỉ thuộc một tài liệu.
2. questions: Phần bài tập CHỈ dựa trên nội dung chương này, gồm đủ 4 dạng: trắc nghiệm một đáp án, tính toán, điền khuyết và kẻ bảng (có thể thêm trắc nghiệm nhiều đáp án và đúng/sai nếu phù hợp). Mỗi câu hỏi có trường type để phân biệt và trường id đánh số từ 1. Với dạng table, trả về table_data gồm headers và rows (dùng ___ cho ô trống).

Số lượng câu hỏi và số câu của từng dạng được quy định chính xác ở phần "PHÂN BỔ CÂU HỎI BẮT BUỘC" bên dưới — phải tuân thủ đúng, không tự ý giảm số câu."""

# Sent as a second (system) message so the JSON contract is explicit while the
# user prompt above stays as specified.
SYSTEM_PROMPT = """Bạn là một gia sư đại học tận tâm phụ trách tổng hợp nhiều tài liệu của cùng một môn học.
Bạn CHỈ trả về một đối tượng JSON hợp lệ, không kèm markdown, không kèm giải thích ngoài JSON.
Tuân thủ đúng cấu trúc sau (các câu hỏi dưới đây chỉ minh hoạ cấu trúc, KHÔNG phải số câu phải tạo):

{
  "title": "Tên bài/chương bao quát nội dung của tất cả tài liệu",
  "teaching": {
    "summary_de": "Zusammenfassung des gesamten Materials auf DEUTSCH (4-8 Sätze), mit Angabe, welches Dokument welchen Punkt ergänzt",
    "summary_vi": "Tóm tắt TỔNG HỢP nội dung của tất cả tài liệu bằng TIẾNG VIỆT (4-8 câu), nêu rõ tài liệu nào bổ sung ý nào",
    "concepts_de": [
      {
        "name": "Begriff auf Deutsch",
        "explanation": "Kurze Erklärung des Begriffs auf Deutsch",
        "sources": ["Dateiname, der den Begriff enthält"]
      }
    ],
    "concepts_vi": [
      {
        "name": "Tên khái niệm",
        "explanation": "Giải thích ngắn gọn khái niệm bằng tiếng Việt",
        "sources": ["tên file chứa khái niệm"]
      }
    ],
    "formulas": [
      {
        "formula": "Công thức, ví dụ: A = B + C",
        "description": "Ý nghĩa của công thức",
        "symbols": [
          {"symbol": "A", "meaning": "Ý nghĩa ký hiệu A"},
          {"symbol": "B", "meaning": "Ý nghĩa ký hiệu B"}
        ],
        "sources": ["tên file chứa công thức"]
      }
    ],
    "examples": [
      {
        "title": "Tên ví dụ 1",
        "problem": "Đề bài / tình huống cụ thể, có thể ghép dữ kiện từ nhiều tài liệu",
        "solution": "Lời giải chi tiết từng bước",
        "sources": ["tên file liên quan"]
      },
      {
        "title": "Tên ví dụ 2",
        "problem": "Đề bài / tình huống cụ thể",
        "solution": "Lời giải chi tiết từng bước",
        "sources": ["tên file liên quan"]
      }
    ],
    "cross_document_links": [
      {
        "documents": ["tài liệu 1.pdf", "tài liệu 2.pdf"],
        "relation": "Kiến thức ở tài liệu 1 là nền tảng để áp dụng công thức ở tài liệu 2"
      }
    ]
  },
  "questions": [
    {
      "id": 1,
      "type": "multiple_choice",
      "question": "Nội dung câu hỏi",
      "options": {"A": "Đáp án A", "B": "Đáp án B", "C": "Đáp án C", "D": "Đáp án D"},
      "answer": "A",
      "explanation": "Giải thích đáp án đúng",
      "difficulty": "Dễ | Trung bình | Khó"
    },
    {
      "id": 2,
      "type": "multiple_response",
      "question": "Chọn TẤT CẢ các đáp án đúng",
      "options": {"A": "Đáp án đúng thứ nhất", "B": "Đáp án sai", "C": "Đáp án đúng thứ hai", "D": "Đáp án sai khác"},
      "answer": ["A", "C"],
      "explanation": "Giải thích vì sao A và C đúng, B và D sai",
      "difficulty": "Dễ | Trung bình | Khó"
    },
    {
      "id": 3,
      "type": "calculation",
      "question": "Đề bài yêu cầu tính toán",
      "given": ["Dữ kiện 1", "Dữ kiện 2"],
      "formula": "Công thức áp dụng (có thể là chuỗi rỗng)",
      "answer": "Kết quả cuối cùng kèm đơn vị, ví dụ: 1.200.000 đồng",
      "explanation": "Các bước tính và giải thích kết quả",
      "difficulty": "Dễ | Trung bình | Khó"
    },
    {
      "id": 4,
      "type": "fill_in_the_blank",
      "question": "Câu hỏi có chỗ trống đánh dấu bằng ___ (mỗi ô trống là một cụm ___ riêng)",
      "answer": ["Đáp án ô trống 1", "Đáp án ô trống 2"],
      "explanation": "Giải thích đáp án đúng",
      "difficulty": "Dễ | Trung bình | Khó"
    },
    {
      "id": 5,
      "type": "true_false",
      "question": "Nhận định cần đánh giá đúng hay sai (ví dụ: Tài sản ngắn hạn luôn có thời hạn trên một năm.)",
      "answer": "false",
      "explanation": "Giải thích vì sao nhận định đúng hoặc sai",
      "difficulty": "Dễ | Trung bình | Khó"
    },
    {
      "id": 6,
      "type": "table",
      "question": "Điền số liệu còn thiếu vào bảng cân đối kế toán sau",
      "table_data": {
        "headers": ["Chỉ tiêu", "Số đầu kỳ", "Số cuối kỳ"],
        "rows": [
          ["Tài sản ngắn hạn", "___", "600"],
          ["Tài sản dài hạn", "400", "___"],
          ["Tổng tài sản", "1.000", "1.200"]
        ]
      },
      "answer": [
        {"row": 0, "column": 1, "value": "600"},
        {"row": 1, "column": 2, "value": "800"}
      ],
      "explanation": "Giải thích cách suy ra số liệu còn thiếu",
      "difficulty": "Dễ | Trung bình | Khó"
    }
  ]
}

Yêu cầu bắt buộc:
- "title" là tên bài/chương bao quát nội dung của TẤT CẢ tài liệu trong môn.
- "teaching" phải có đủ "summary_de", "summary_vi", "concepts_de", "concepts_vi", "formulas", "examples", "cross_document_links".
- "summary_de" và "summary_vi" phải TỔNG HỢP nội dung của mọi tài liệu được cung cấp (không tóm tắt riêng lẻ từng file) và nêu rõ tài liệu nào bổ sung ý nào; summary_de viết bằng tiếng Đức, summary_vi viết bằng tiếng Việt.
- "concepts_de" và "concepts_vi" là hai mảng khái niệm CÙNG số lượng, CÙNG thứ tự; mỗi khái niệm có "name", "explanation" ngắn gọn và "sources" (mảng tên file chứa khái niệm); concepts_de bằng tiếng Đức, concepts_vi bằng tiếng Việt.
- "formulas" là mảng đối tượng; mỗi công thức có "formula", "description", "symbols" (mỗi ký hiệu có "symbol" và "meaning") và "sources". Nếu tài liệu không có công thức thì trả về mảng rỗng [].
- "examples" phải có 2-3 ví dụ minh hoạ cụ thể, mỗi ví dụ có "title", "problem", "solution" là lời giải chi tiết và "sources".
- "cross_document_links" phải có ít nhất 2 mục khi môn học có từ 2 tài liệu trở lên; mỗi mục nêu "documents" (tên các file liên quan) và "relation" (mối liên hệ kiến thức).
- Số lượng câu hỏi và số câu của từng dạng do phần "PHÂN BỔ CÂU HỎI BẮT BUỘC" quy định; "questions" phải có ĐÚNG số câu đó, "id" chạy liên tục từ 1, không trùng.
- Mỗi câu hỏi phải có "type" đúng dạng đã phân bổ, cùng "explanation" và "difficulty" ("Dễ", "Trung bình" hoặc "Khó").
- multiple_choice: đúng 4 lựa chọn A, B, C, D; "answer" là một chữ cái (A, B, C hoặc D).
- multiple_response: 4-5 lựa chọn A, B, C, D (thêm E nếu cần); có từ 2 đáp án đúng trở lên; "answer" là mảng các chữ cái đúng, ví dụ ["A", "C"].
- true_false: "question" là một nhận định rõ ràng; "answer" là "true" hoặc "false" (chuỗi, không kèm giải thích trong trường này).
- calculation: nêu rõ "given" (dữ kiện) và "formula" (công thức, nếu không có thì để chuỗi rỗng); "answer" là kết quả cuối cùng, nêu kèm đơn vị.
- fill_in_the_blank: mỗi ô trống viết là "___"; "answer" là mảng đáp án đúng theo đúng thứ tự các ô trống trong câu hỏi (số phần tử bằng số ô "___").
- table: LUÔN có "table_data" gồm "headers" (mảng tiêu đề cột) và "rows" (mảng các hàng, mỗi hàng là mảng chuỗi có cùng độ dài với "headers"); ô cần điền đánh dấu bằng "___"; "answer" liệt kê từng ô cần điền với "row", "column" là chỉ số hàng/cột bắt đầu từ 0 (theo "rows" và "headers") và "value" là giá trị đúng.
- Khi môn học có nhiều tài liệu, câu hỏi phải khuyến khích liên kết kiến thức giữa các tài liệu (dùng dữ kiện của tài liệu này để áp dụng kiến thức ở tài liệu kia).
- Nội dung phải bám sát tài liệu được cung cấp, không bịa thêm kiến thức ngoài tài liệu.
"""

# --------------------------------------------------------------------------------------
# Output language
# --------------------------------------------------------------------------------------
# The web client sends the UI locale (`lang`) so the stored curriculum can be
# displayed again in the same language. Unknown/missing values fall back to
# Vietnamese, preserving the original behaviour of this service.
# ``LANGUAGE_DIRECTIVES`` dùng cho /generate-exam (chỉ sinh câu hỏi);
# /process-document dùng ``DOCUMENT_LANGUAGE_DIRECTIVES`` vì ``teaching`` là
# song ngữ Đức - Việt bất kể ``lang``.
LANGUAGE_DIRECTIVES = {
    "vi": (
        "Ngôn ngữ đầu ra: TIẾNG VIỆT. Viết toàn bộ questions, options, answer "
        "và explanation bằng tiếng Việt; giữ nguyên công thức và ký hiệu toán học."
    ),
    "de": (
        "Ngôn ngữ đầu ra: TIẾNG ĐỨC (Deutsch). Schreibe questions, options, "
        "answer und explanation vollständig auf Deutsch; Formeln und "
        "mathematische Symbole bleiben unverändert."
    ),
}

DEFAULT_OUTPUT_LANGUAGE = "vi"

# Ngôn ngữ của các trường song ngữ trong teaching / 4 phần của khái niệm.
# Thứ tự cố định để output luôn ổn định: tiếng Đức trước, tiếng Việt sau.
BILINGUAL_LANGUAGES: tuple[str, ...] = ("de", "vi")

# Khoá ngôn ngữ model có thể dùng (Deutsch, Vietnamese, vi-VN, ...).
_LANGUAGE_KEY_ALIASES: dict[str, frozenset[str]] = {
    "de": frozenset({"de", "deu", "ger", "german", "deutsch", "tieng_duc"}),
    "vi": frozenset({"vi", "vie", "vn", "viet", "vietnamese", "tieng_viet", "tiengviet"}),
}

DOCUMENT_LANGUAGE_DIRECTIVES = {
    "vi": (
        "Ngôn ngữ đầu ra: TIẾNG VIỆT cho formulas (description và meaning), "
        "examples, cross_document_links, questions, options, answer và explanation; "
        "giữ nguyên công thức và ký hiệu toán học. Riêng teaching là SONG NGỮ: "
        "summary_vi/concepts_vi bằng tiếng Việt, summary_de/concepts_de bằng tiếng Đức."
    ),
    "de": (
        "Ngôn ngữ đầu ra: TIẾNG ĐỨC (Deutsch) cho formulas (description und meaning), "
        "examples, cross_document_links, questions, options, answer und explanation; "
        "Formeln und mathematische Symbole bleiben unverändert. Das Feld teaching ist "
        "ZWEISPRACHIG: summary_de/concepts_de auf Deutsch, summary_vi/concepts_vi auf "
        "Vietnamesisch."
    ),
}

BILINGUAL_TEACHING_DIRECTIVE = """YÊU CẦU SONG NGỮ ĐỨC - VIỆT CHO teaching (bắt buộc):
- teaching.summary_de: bản tóm tắt bằng TIẾNG ĐỨC; teaching.summary_vi: cùng nội dung bằng TIẾNG VIỆT. Hai bản phải tương đương về ý và cùng độ chi tiết.
- teaching.concepts_de và teaching.concepts_vi: hai mảng khái niệm CÙNG số lượng, CÙNG thứ tự, CÙNG nội dung; concepts_de viết bằng tiếng Đức, concepts_vi viết bằng tiếng Việt, mỗi khái niệm giữ nguyên "sources".
- Tên riêng, tên tài liệu, công thức và ký hiệu toán học giữ nguyên ở cả hai bản, không dịch.
- KHÔNG trả về "summary" hay "key_concepts" không hậu tố; chỉ dùng đúng bốn khoá summary_de, summary_vi, concepts_de, concepts_vi."""


def resolve_language(lang: str | None) -> str:
    """Map a client language tag ("de", "vi", "de-DE", ...) to a supported one."""
    normalized = (lang or "").strip().lower().replace("_", "-")
    short = normalized.split("-", 1)[0]
    return short if short in LANGUAGE_DIRECTIVES else DEFAULT_OUTPUT_LANGUAGE


CROSS_DOCUMENT_DIRECTIVE = """YÊU CẦU TỔNG HỢP NHIỀU TÀI LIỆU (bắt buộc):
- teaching.summary_de và teaching.summary_vi phải nối kiến thức của TẤT CẢ tài liệu trong môn; nêu rõ tài liệu nào bổ sung hoặc đối chiếu ý nào.
- teaching.cross_document_links phải có ÍT NHẤT 2 mục, mỗi mục là một cặp tài liệu có liên hệ kiến thức thật sự.
- concepts_de, concepts_vi, formulas, examples phải ghi "sources" là tên file tương ứng.
- Câu hỏi phải kiểm tra khả năng liên kết kiến thức giữa các tài liệu khi nội dung cho phép."""

# Dùng cho chế độ mặc định ``per_file``: mỗi file được phân tích độc lập, nên
# model KHÔNG được trộn nội dung của các file khác vào kết quả của file này.
SINGLE_DOCUMENT_DIRECTIVE = """YÊU CẦU PHÂN TÍCH ĐỘC LẬP (bắt buộc):
- Chỉ phân tích duy nhất tài liệu này. Không gộp nội dung với bất kỳ tài liệu nào khác. Tạo tóm tắt và khái niệm riêng cho chương này.
- Mọi trường của teaching (summary_de, summary_vi, concepts_de, concepts_vi, formulas, examples) và toàn bộ questions chỉ được lấy từ nội dung của tài liệu đang phân tích; không nhắc tới, không suy diễn và không bổ sung kiến thức từ file khác.
- Nếu tài liệu không có nội dung cho một mục, trả về mảng rỗng cho mục đó thay vì lấy từ tài liệu khác.
- teaching.cross_document_links luôn là mảng rỗng [] vì kết quả này chỉ thuộc một tài liệu."""


def build_system_prompt(
    lang: str | None,
    count: int | None = None,
    plan: Sequence[str] | None = None,
    cross_document: bool = False,
    single_document: bool = False,
) -> str:
    """SYSTEM_PROMPT + question plan + cross/single-document rules + language directives.

    ``teaching`` luôn có bản ``de`` và ``vi`` (``BILINGUAL_TEACHING_DIRECTIVE``);
    các trường còn lại theo ``lang`` của client. Backwards compatible:
    ``build_system_prompt("de")`` vẫn trả về contract gốc kèm chỉ dẫn ngôn ngữ.

    ``single_document=True`` (mode ``per_file``) thêm ``SINGLE_DOCUMENT_DIRECTIVE``
    và bỏ luật cross-document: mỗi file phải được phân tích độc lập.
    """
    parts = [SYSTEM_PROMPT]
    if plan:
        parts.append(build_question_plan_block(plan))
    elif count:
        parts.append(build_question_plan_block(plan_question_types(count)))
    if single_document:
        parts.append(SINGLE_DOCUMENT_DIRECTIVE)
    elif cross_document:
        parts.append(CROSS_DOCUMENT_DIRECTIVE)
    parts.append(BILINGUAL_TEACHING_DIRECTIVE)
    parts.append(DOCUMENT_LANGUAGE_DIRECTIVES[resolve_language(lang)])
    return "\n\n".join(parts)


def build_document_user_prompt(
    *,
    corpus: str,
    plan: Sequence[str],
    subject: str = "",
    documents: Sequence["ExtractedDocument"] = (),
    focus_filename: str | None = None,
    single_document: bool = False,
) -> str:
    """User message for /process-document: task + plan + the document(s) sent.

    ``single_document=True`` (mode ``per_file``) chỉ nhúng nội dung của đúng file
    đang xử lý và nhấn mạnh yêu cầu phân tích độc lập, không liệt kê các file khác
    của môn. ``combined`` vẫn giữ nguyên hành vi cũ (toàn bộ tài liệu của môn).
    """
    lines = [
        SINGLE_DOCUMENT_USER_PROMPT if single_document else USER_PROMPT,
        "",
        build_question_plan_block(plan),
    ]
    if subject.strip():
        lines.append(f"Môn học: {subject.strip()}")
    if single_document:
        if focus_filename:
            lines.append(f"Tài liệu duy nhất cần phân tích: {focus_filename}")
        lines.append("")
        lines.append(SINGLE_DOCUMENT_DIRECTIVE)
    else:
        if focus_filename:
            lines.append(f"Tài liệu chính của phần này: {focus_filename}")
        if documents:
            lines.append("")
            lines.append("===== DANH SÁCH TÀI LIỆU TRONG MÔN =====")
            lines.extend(
                f"- TÀI LIỆU {doc.position}: {doc.filename} ({doc.pages} trang, "
                f"{doc.characters} ký tự{', đã cắt bớt' if doc.truncated else ''})"
                for doc in documents
            )
    lines.append("")
    lines.append(
        "===== NỘI DUNG TÀI LIỆU DUY NHẤT ====="
        if single_document
        else "===== TOÀN BỘ NỘI DUNG TÀI LIỆU ====="
    )
    lines.append(corpus)
    return "\n".join(lines)


def build_topup_prompt(
    *,
    shortfall: int,
    plan: Sequence[str],
    generated: Sequence[dict[str, Any]],
    corpus: str,
    subject: str = "",
    focus_filename: str | None = None,
    context_label: str = "TOÀN BỘ NỘI DUNG TÀI LIỆU",
    directive: str | None = None,
) -> str:
    """Ask the model for exactly the questions missing from a short answer."""
    aligned = align_questions_to_plan(generated, plan)
    deficits = question_type_deficits(aligned, plan)
    lines = [
        "===== YÊU CẦU BỔ SUNG CÂU HỎI =====",
        (
            f"Đề hiện có {len(aligned)}/{len(plan)} câu. "
            f"Hãy tạo THÊM ĐÚNG {shortfall} câu hỏi MỚI."
        ),
        "- Số câu còn thiếu theo từng dạng: "
        + "; ".join(f"{qtype}={amount}" for qtype, amount in deficits),
        '- Chỉ trả về JSON dạng {"questions": [...]} gồm đúng số câu trên; "id" đánh số từ 1.',
        "- KHÔNG lặp lại nội dung của những câu đã có:",
    ]
    lines.extend(f"- {stem}" for stem in question_stems(generated))
    if directive:
        lines.append("")
        lines.append(directive)
    if subject.strip():
        lines.append(f"Môn học: {subject.strip()}")
    if focus_filename:
        lines.append(f"Tài liệu chính: {focus_filename}")
    lines.append("")
    lines.append(f"===== {context_label} =====")
    lines.append(corpus)
    return "\n".join(lines)


# --------------------------------------------------------------------------------------
# Exam generation prompt (POST /generate-exam)
# --------------------------------------------------------------------------------------
# The exam endpoint receives a compact "knowledge framework" (summary, key
# concepts, formulas, worked examples) instead of the raw document text. That is enough
# to write new questions while keeping the prompt small and cheap.
EXAM_SYSTEM_PROMPT = """Bạn là chuyên gia khảo thí của một trường đại học.
Bạn CHỈ trả về một đối tượng JSON hợp lệ, không kèm markdown, không kèm giải thích ngoài JSON.

Cấu trúc bắt buộc:
{
  "questions": [
    {
      "id": 1,
      "type": "multiple_choice",
      "question": "Nội dung câu hỏi",
      "options": {"A": "Đáp án A", "B": "Đáp án B", "C": "Đáp án C", "D": "Đáp án D"},
      "answer": "B",
      "explanation": "Giải thích đáp án đúng",
      "difficulty": "Dễ | Trung bình | Khó"
    }
  ]
}

Sáu dạng câu hỏi được phép (trường "type"):
- "multiple_choice": đúng 4 lựa chọn A, B, C, D; "answer" là một chữ cái.
- "multiple_response": 4-5 lựa chọn; có từ 2 đáp án đúng; "answer" là mảng chữ cái, ví dụ ["A", "C"].
- "calculation": có "given" (mảng dữ kiện), "formula" (chuỗi, có thể rỗng) và "answer" là kết quả kèm đơn vị.
- "fill_in_the_blank": mỗi chỗ trống viết là "___"; "answer" là mảng đáp án theo đúng thứ tự các ô trống.
- "true_false": "question" là một nhận định rõ ràng; "answer" là "true" hoặc "false".
- "table": có "table_data": {"headers": [...], "rows": [[...]]} với ô cần điền là "___";
  "answer" là mảng {"row": số_hàng, "column": số_cột, "value": "giá trị đúng"} (chỉ số bắt đầu từ 0).

Yêu cầu bắt buộc:
- Tạo ĐÚNG số lượng câu hỏi được yêu cầu, "id" chạy từ 1 đến hết, không trùng.
- Mỗi câu có đủ "type", "question", "answer", "explanation", "difficulty" ("Dễ", "Trung bình" hoặc "Khó").
- Bốn dạng bắt buộc: trắc nghiệm một đáp án ("multiple_choice"), tính toán ("calculation"), điền khuyết ("fill_in_the_blank") và kẻ bảng ("table"); có thể thêm "multiple_response" và "true_false". Số câu của từng dạng do phần "PHÂN BỔ CÂU HỎI BẮT BUỘC" quy định.
- Câu hỏi phải MỚI: không sao chép câu hỏi/đáp án đã có, không lặp lại cùng một câu hỏi với số liệu khác ở hai câu liền nhau.
- Đáp án đúng nằm ở vị trí ngẫu nhiên giữa A, B, C, D; không luôn đặt ở A.
- Độ khó phân bổ hợp lý (khoảng 30% Dễ, 50% Trung bình, 20% Khó) và thay đổi giữa các lần tạo đề.
- Nội dung bám sát khuôn khổ kiến thức được cung cấp, không bịa thêm kiến thức ngoài tài liệu.
- Mọi câu hỏi phải chấm được tự động: đáp án trùng khớp với options/ô trống/bảng đã nêu.
- Khi đề bao gồm nhiều chương, phân bổ câu hỏi cho TẤT CẢ các chương, không bỏ sót chương nào."""

EXAM_SCOPE_LABELS = {
    "chapter": "Đề luyện tập cho MỘT chương của môn học",
    "final": "Đề thi kết thúc học phần, tổng hợp TẤT CẢ các chương",
}


class ExamRequest(BaseModel):
    """Payload of ``POST /generate-exam`` (built by the Next.js API route)."""

    context: str = Field(
        description=(
            "Khuôn khổ kiến thức: tóm tắt, khái niệm chính, công thức, ví dụ. "
            "Nhiều chương được ngăn cách bằng dòng '### CHƯƠNG …'."
        )
    )
    count: int = Field(
        default=10,
        description="Số câu hỏi muốn tạo (được kẹp trong EXAM_MIN/MAX_QUESTIONS)",
    )
    types: list[str] = Field(
        default_factory=list,
        description=(
            "Các dạng câu hỏi muốn có (multiple_choice, calculation, "
            "fill_in_the_blank, table, ...). Bỏ trống = cả 4 dạng bắt buộc."
        ),
    )
    scope: str = Field(
        default="chapter", description="'chapter' hoặc 'final' (quyết định cách mô tả đề)"
    )
    lang: str = Field(
        default=DEFAULT_OUTPUT_LANGUAGE,
        description="Ngôn ngữ đầu ra: 'de' hoặc 'vi' (mặc định: vi)",
    )
    chapter_titles: list[str] = Field(
        default_factory=list,
        description="Tiêu đề các chương nằm trong phạm vi đề (dùng để phân bổ câu hỏi)",
    )
    existing_questions: list[str] = Field(
        default_factory=list,
        description="Các câu hỏi đã có, model phải tránh lặp lại (tối đa 15 câu)",
    )
    seed: int = Field(
        default=0,
        description="Số ngẫu nhiên do client sinh ra cho mỗi lần tạo đề",
    )


def clamp_exam_count(count: int) -> int:
    """Keep the requested question count inside the configured bounds."""
    return max(EXAM_MIN_QUESTIONS, min(EXAM_MAX_QUESTIONS, int(count)))


def build_exam_prompt(
    request: ExamRequest, count: int, plan: Sequence[str] | None = None
) -> str:
    """User message for exam generation: scope, chapters, plan, framework, variety."""
    scope = request.scope if request.scope in EXAM_SCOPE_LABELS else "chapter"
    titles = [title.strip() for title in request.chapter_titles if title.strip()]
    existing = [
        question.strip()
        for question in request.existing_questions
        if question.strip()
    ][:MAX_EXAM_EXISTING_QUESTIONS]
    question_plan = list(plan) if plan else plan_question_types(count, request.types)

    lines = [
        f"Nhiệm vụ: tạo một đề luyện thi MỚI gồm ĐÚNG {count} câu hỏi.",
        f"Phạm vi: {EXAM_SCOPE_LABELS[scope]}.",
    ]

    if titles:
        lines.append("Các chương trong phạm vi (phân bổ câu hỏi cho từng chương):")
        lines.extend(f"- {title}" for title in titles)

    if existing:
        lines.append(
            "KHÔNG được lặp lại (hoặc chỉ đổi số liệu của) những câu hỏi đã có sau đây:"
        )
        lines.extend(f"- {question}" for question in existing)

    lines.append(
        f"Mã biến thể: {request.seed}. Hãy dùng mã này để tạo một đề khác hoàn toàn "
        "so với lần trước: đổi góc hỏi, đổi số liệu, đổi thứ tự đáp án."
    )
    if question_plan:
        lines.append("")
        lines.append(build_question_plan_block(question_plan))
    lines.append("")
    lines.append("===== KHUÔN KHỔ KIẾN THỨC =====")
    lines.append(request.context.strip()[:MAX_EXAM_CONTEXT_CHARS])

    return "\n".join(lines)


def build_exam_system_prompt(
    lang: str | None, plan: Sequence[str] | None = None
) -> str:
    """EXAM_SYSTEM_PROMPT + question plan + output-language directive."""
    parts = [EXAM_SYSTEM_PROMPT]
    if plan:
        parts.append(build_question_plan_block(plan))
    parts.append(LANGUAGE_DIRECTIVES[resolve_language(lang)])
    return "\n\n".join(parts)


# --------------------------------------------------------------------------------------
# Concept explanation prompt (POST /explain-concept)
# --------------------------------------------------------------------------------------
# Bốn khoá của hợp đồng JSON là tiếng Đức vì client StudyGap hiển thị nguyên tên khoá:
# definition / erklaerung / zusammenhang / kritik. Mỗi khoá chứa object song ngữ
# {"de": ..., "vi": ...} nên client có ngay cả bản tiếng Đức và tiếng Việt; tên khoá
# không bao giờ bị dịch.
CONCEPT_SECTIONS: tuple[str, ...] = (
    "definition",
    "erklaerung",
    "zusammenhang",
    "kritik",
)

CONCEPT_SYSTEM_PROMPT = """Bạn là một gia sư đại học tận tâm, chuyên giải thích khái niệm của môn học cho sinh viên.
Bạn CHỈ trả về một đối tượng JSON hợp lệ, không kèm markdown, không kèm giải thích ngoài JSON.

Cấu trúc bắt buộc (giữ nguyên tên khoá, đúng thứ tự; MỖI phần là object song ngữ với đúng hai khoá "de" và "vi"):
{
  "definition": {
    "de": "Knappe, präzise Definition auf Deutsch (1-3 Sätze).",
    "vi": "Định nghĩa ngắn gọn, chính xác bằng tiếng Việt (1-3 câu)."
  },
  "erklaerung": {
    "de": "Ausführliche, leicht verständliche Erklärung auf Deutsch mit mindestens einem konkreten Beispiel.",
    "vi": "Giải thích chi tiết, dễ hiểu bằng tiếng Việt, có ít nhất một ví dụ minh hoạ cụ thể."
  },
  "zusammenhang": {
    "de": "Bezug des Begriffs zu anderen Begriffen des Fachs auf Deutsch.",
    "vi": "Mối liên hệ của khái niệm với các khái niệm khác trong môn học, bằng tiếng Việt."
  },
  "kritik": {
    "de": "Kritik, Vor- und Nachteile oder offene Diskussionen auf Deutsch; leerer String, wenn es nichts zu kritisieren gibt.",
    "vi": "Phê bình, ưu nhược điểm hoặc các tranh luận xung quanh khái niệm bằng tiếng Việt; chuỗi rỗng nếu không có."
  }
}

Yêu cầu bắt buộc:
- Mỗi phần PHẢI là object có đúng hai khoá "de" và "vi"; không trả về chuỗi trần, không thêm khoá khác.
- Bản "de" viết hoàn toàn bằng tiếng Đức, bản "vi" viết hoàn toàn bằng tiếng Việt; hai bản cùng nội dung và cùng độ chi tiết.
- definition: CHỈ là định nghĩa, 1-3 câu, chính xác theo tài liệu; không kể lể, không lặp nguyên văn cả đoạn dài.
- erklaerung: giải thích chi tiết bằng ngôn ngữ dễ hiểu cho sinh viên, có ÍT NHẤT MỘT ví dụ minh hoạ cụ thể (số liệu hoặc tình huống thực tế nếu phù hợp).
- zusammenhang: nêu rõ khái niệm liên hệ với những khái niệm nào khác của môn học và kiểu quan hệ (nền tảng, mở rộng, đối lập, công cụ tính toán, ...).
- kritik: nêu ưu điểm, nhược điểm, giới hạn áp dụng hoặc các tranh luận xung quanh khái niệm (nếu có); nếu không có gì đáng phê bình thì trả về chuỗi rỗng cho cả hai ngôn ngữ.
- Tên riêng, thuật ngữ chuyên ngành, công thức và ký hiệu toán học giữ nguyên ở cả hai bản, không dịch.
- Nội dung bám sát tài liệu được cung cấp, không bịa thêm kiến thức ngoài tài liệu. Nếu tài liệu không đề cập trực tiếp khái niệm, hãy dùng kiến thức chuẩn của môn học và ghi rõ điều đó trong kritik (cả hai ngôn ngữ)."""

# Language directive riêng: LANGUAGE_DIRECTIVES liệt kê các trường của
# teaching/questions nên không dùng lại được cho 4 phần của khái niệm.
# ``lang`` chỉ quyết định ngôn ngữ hiển thị chính; hợp đồng song ngữ luôn bắt buộc.
CONCEPT_LANGUAGE_DIRECTIVES = {
    "vi": (
        "Ngôn ngữ hiển thị chính: TIẾNG VIỆT (client gửi lang='vi'). Hợp đồng song ngữ "
        "vẫn bắt buộc: mọi phần luôn có đủ bản \"de\" (tiếng Đức) và \"vi\" (tiếng Việt); "
        "giữ nguyên tên khoá JSON, công thức và ký hiệu toán học."
    ),
    "de": (
        "Haupt-Anzeigesprache: Deutsch (client sendet lang='de'). Der zweisprachige "
        "Vertrag ist trotzdem Pflicht: jede Sektion enthält immer sowohl \"de\" "
        "(Deutsch) als auch \"vi\" (Vietnamesisch); JSON-Schlüssel, Formeln und "
        "mathematische Symbole bleiben unverändert."
    ),
}


class ConceptRequest(BaseModel):
    """Payload of ``POST /explain-concept``."""

    concept: str = Field(description="Tên khái niệm cần phân tích")
    context: str = Field(
        default="",
        description=(
            "Nội dung tài liệu liên quan (client ghép sẵn). Được cắt theo "
            "MAX_CONCEPT_CONTEXT_CHARS; bỏ trống thì model dựa vào kiến thức chuẩn."
        ),
    )
    lang: str = Field(
        default=DEFAULT_OUTPUT_LANGUAGE,
        description=(
            "Ngôn ngữ hiển thị chính: 'de' hoặc 'vi' (mặc định: vi). Mọi phần luôn "
            "trả về cả bản 'de' và 'vi'."
        ),
    )


def build_concept_system_prompt(lang: str | None) -> str:
    """CONCEPT_SYSTEM_PROMPT + bilingual contract + display-language directive."""
    directive = CONCEPT_LANGUAGE_DIRECTIVES[resolve_language(lang)]
    return "\n\n".join([CONCEPT_SYSTEM_PROMPT, directive])


def build_concept_user_prompt(*, concept: str, context: str = "") -> str:
    """User message of /explain-concept: the concept + its related document text."""
    context = context.strip()
    lines = [
        f'Nhiệm vụ: phân tích khái niệm "{concept}" của môn học theo đúng 4 phần sau:',
        "1. definition: Định nghĩa ngắn gọn, chính xác.",
        "2. erklaerung: Giải thích chi tiết, dễ hiểu, có ví dụ minh hoạ.",
        "3. zusammenhang: Mối liên hệ của khái niệm với các khái niệm khác trong môn học.",
        "4. kritik: Phê bình, ưu nhược điểm hoặc các tranh luận xung quanh khái niệm (nếu có).",
        'Mỗi phần trả về object song ngữ {"de": "...", "vi": "..."}: bản "de" viết bằng tiếng Đức, bản "vi" viết bằng tiếng Việt.',
        "",
        "===== KHÁI NIỆM CẦN PHÂN TÍCH =====",
        concept,
        "",
        "===== NỘI DUNG TÀI LIỆU LIÊN QUAN =====",
        context
        or (
            "(Không có tài liệu kèm theo: hãy dùng kiến thức chuẩn của môn học "
            "và ghi rõ điều đó trong kritik.)"
        ),
    ]
    return "\n".join(lines)


# Model có thể trả khoá "Erklärung" (có dấu), camelCase hoặc tiếng Anh/tiếng Việt,
# nên mỗi phần chấp nhận nhiều alias; giá trị có thể là chuỗi, mảng hoặc object.
_CONCEPT_SECTION_ALIASES: dict[str, frozenset[str]] = {
    "definition": frozenset(
        {"definition", "begriffsdefinition", "definierung", "dinh_nghia"}
    ),
    "erklaerung": frozenset(
        {"erklaerung", "erklarung", "explanation", "detailed_explanation", "giai_thich"}
    ),
    "zusammenhang": frozenset(
        {
            "zusammenhang",
            "kontext",
            "beziehung",
            "beziehungen",
            "relation",
            "moi_lien_he",
            "lien_he",
        }
    ),
    "kritik": frozenset(
        {
            "kritik",
            "kritische_wuerdigung",
            "critique",
            "criticism",
            "phe_binh",
            "uu_nhuoc_diem",
        }
    ),
}


def _section_key(value: Any) -> str:
    """Normalize a JSON key to ``[a-z0-9_]`` (umlauts spelled out: ä -> ae)."""
    text = str(value).strip().lower()
    for umlaut, replacement in (("ä", "ae"), ("ö", "oe"), ("ü", "ue"), ("ß", "ss")):
        text = text.replace(umlaut, replacement)
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


def _section_text(value: Any) -> str:
    """Flatten a section value (string / list / object) into plain text."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, dict):
        lines = []
        for key, item in value.items():
            text = _section_text(item)
            if text:
                lines.append(f"{key}: {text}")
        return "\n".join(lines)
    if isinstance(value, (list, tuple)):
        return "\n".join(
            text for text in (_section_text(item) for item in value) if text
        )
    return ""


def _new_bilingual_text() -> dict[str, str]:
    """A fresh ``{"de": "", "vi": ""}`` object (không dùng chung dict giữa các phần)."""
    return {code: "" for code in BILINGUAL_LANGUAGES}


def language_code(value: Any) -> str | None:
    """Nhận diện khoá ngôn ngữ: ``de``, ``DE``, ``Deutsch``, ``vi-VN``, ``Tiếng Việt``."""
    key = _section_key(value)
    for code, aliases in _LANGUAGE_KEY_ALIASES.items():
        if key in aliases:
            return code
    return None


def localized_text(value: Any, *, lang: str = DEFAULT_OUTPUT_LANGUAGE) -> dict[str, str]:
    """Chuẩn hoá một giá trị thành object song ngữ ``{"de": ..., "vi": ...}``.

    - Object ``{"de": ..., "vi": ...}`` (kể cả alias ``Deutsch``/``Vietnamese``)
      được giữ nguyên cả hai bản.
    - Chuỗi/mảng/object thường được coi là bản của ``lang`` (tương thích hợp đồng
      cũ: model chỉ trả một ngôn ngữ thì client vẫn nhận được nội dung).
    """
    result = _new_bilingual_text()
    if isinstance(value, dict):
        for key, item in value.items():
            code = language_code(key)
            if code and not result[code]:
                result[code] = _section_text(item)
        if any(result.values()):
            return result
    text = _section_text(value)
    if text:
        result[resolve_language(lang)] = text
    return result


def normalize_concept_explanation(
    payload: Any, *, lang: str = DEFAULT_OUTPUT_LANGUAGE
) -> dict[str, dict[str, str]]:
    """Trả đúng 4 phần, mỗi phần là object song ngữ ``{"de": str, "vi": str}``.

    Tolerant như ``normalize_teaching``: "Erklärung" có dấu, camelCase hoặc alias
    tiếng Anh/tiếng Việt, khoá phẳng ``definition_de``/``definition_vi``, giá trị
    mảng/object và một lớp bọc ngoài (``{"concept": {...}}``) đều được chấp nhận.
    Chuỗi trần được coi là bản của ``lang`` nên model trả hợp đồng cũ vẫn không
    làm client lỗi.
    """
    result = {section: _new_bilingual_text() for section in CONCEPT_SECTIONS}
    if not isinstance(payload, dict):
        return result

    candidates: list[dict[str, Any]] = [payload]
    candidates.extend(item for item in payload.values() if isinstance(item, dict))

    for candidate in candidates:
        for key, value in candidate.items():
            normalized_key = _section_key(key)
            for section in CONCEPT_SECTIONS:
                aliases = _CONCEPT_SECTION_ALIASES[section]
                matched_lang: str | None = None
                if normalized_key not in aliases:
                    for code in BILINGUAL_LANGUAGES:
                        if normalized_key in {f"{alias}_{code}" for alias in aliases}:
                            matched_lang = code
                            break
                    else:
                        continue
                if matched_lang:
                    # Khoá đã ghi rõ ngôn ngữ ("definition_de") -> chỉ lấy bản đó.
                    texts = localized_text(value, lang=lang)
                    text = texts[matched_lang] or _section_text(value)
                    if text and not result[section][matched_lang]:
                        result[section][matched_lang] = text
                    continue
                for code, text in localized_text(value, lang=lang).items():
                    if text and not result[section][code]:
                        result[section][code] = text

    # Model chỉ trả một khối văn bản duy nhất -> coi đó là phần giải thích.
    if not any(text for texts in result.values() for text in texts.values()):
        fallback = _section_text(payload.get("text")) or _section_text(
            payload.get("answer")
        )
        if fallback:
            result["erklaerung"][resolve_language(lang)] = fallback
    return result


# --------------------------------------------------------------------------------------
# App + CORS
# --------------------------------------------------------------------------------------
app = FastAPI(
    title="StudyGap AI - Document Processing API",
    description=(
        "Chuyển tài liệu thành nội dung giảng dạy SONG NGỮ ĐỨC - VIỆT "
        "(`teaching`: `summary_de`/`summary_vi`, "
        "`concepts_de`/`concepts_vi`, công thức, ví dụ, liên kết giữa các tài liệu) "
        "và bộ câu hỏi ôn tập (`questions`) gồm 4 dạng bắt buộc: trắc nghiệm một đáp án, "
        "tính toán, điền khuyết, kẻ bảng (kèm 2 dạng mở rộng trắc nghiệm nhiều đáp án và "
        "đúng/sai). Số câu do client yêu cầu luôn được bảo đảm đủ (tự bổ sung nếu model "
        "trả thiếu). Định dạng nhận được: PDF, Word (.docx), Excel (.xlsx), text (.txt), "
        "CSV (.csv). Gửi nhiều file trong cùng một request qua field `files`: mặc định "
        "(`mode=per_file`) mỗi file được trích xuất và phân tích ĐỘC LẬP, trả về một phần "
        "tử riêng; `mode=combined` trả về một phần tử tổng hợp cho cả môn.\n\n"
        "`POST /generate-exam` tạo đề luyện tập/đề thi mới từ khuôn khổ kiến thức đã lưu. "
        "`POST /explain-concept` phân tích một khái niệm thành 4 phần: definition "
        "(định nghĩa), erklaerung (giải thích kèm ví dụ), zusammenhang (liên hệ với "
        "khái niệm khác) và kritik (phê bình, ưu nhược điểm)."
    ),
    version="2.1.0",
)

app.add_middleware(
    CORSMiddleware,
    # Allow All Origins: mọi origin đều gọi được API (mặc định "*").
    # Khi allow_credentials=True, Starlette tự phản chiếu lại Origin cụ thể trong
    # Access-Control-Allow-Origin thay vì gửi "*", nên trình duyệt vẫn chấp nhận.
    allow_origins=ALLOWED_ORIGINS,
    # Cho phép domain Vercel bằng regex (hỗ trợ cả preview deployment).
    allow_origin_regex=CORS_ALLOW_ORIGIN_REGEX or None,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
    max_age=600,
)
if CORS_ALLOW_ALL_ORIGINS:
    logger.info("CORS enabled: allow all origins (*)")
else:
    logger.info("CORS enabled for origins: %s", ", ".join(ALLOWED_ORIGINS))
if CORS_ALLOW_ORIGIN_REGEX:
    logger.info("CORS enabled for origin regex (Vercel): %s", CORS_ALLOW_ORIGIN_REGEX)


# --------------------------------------------------------------------------------------
# DeepSeek client (lazily created so the app still boots without an API key)
# --------------------------------------------------------------------------------------
@lru_cache(maxsize=1)
def get_deepseek_client() -> OpenAI:
    """Return a cached OpenAI client pointed at the DeepSeek API.

    API key được đọc từ biến môi trường ``DEEPSEEK_API_KEY`` (hoặc file ``.env``)
    tại thời điểm gọi. Thiếu key thì trả 500 kèm hướng dẫn, không hardcode key
    trong source. Client được cache nên chỉ khởi tạo một lần cho mỗi process.
    """
    api_key = os.getenv("DEEPSEEK_API_KEY", "").strip()
    if not api_key:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=(
                "Thiếu biến môi trường DEEPSEEK_API_KEY. "
                'Đặt bằng lệnh: $env:DEEPSEEK_API_KEY = "sk-..." '
                "hoặc thêm DEEPSEEK_API_KEY vào file .env rồi khởi động lại server."
            ),
        )
    return OpenAI(
        api_key=api_key,
        base_url=DEEPSEEK_BASE_URL,
        timeout=DEEPSEEK_TIMEOUT_SECONDS,
        max_retries=2,
    )


# --------------------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------------------
# Mọi định dạng dưới đây được đọc bằng thư viện riêng rồi chuyển thành plain text,
# nên phần còn lại của pipeline (corpus, prompt, câu hỏi) không cần biết định dạng
# gốc. Muốn hỗ trợ thêm định dạng: viết một hàm ``extract_*_text`` rồi khai báo
# đuôi file trong ``FORMAT_LABELS`` và một nhánh trong ``extract_text``.
FORMAT_LABELS: dict[str, str] = {
    ".pdf": "PDF",
    ".docx": "Word (.docx)",
    ".xlsx": "Excel (.xlsx)",
    ".txt": "văn bản (.txt)",
    ".csv": "CSV (.csv)",
}
# Đuôi "anh em" đọc bằng cùng thư viện với đuôi chuẩn ở trên.
EXTENSION_ALIASES: dict[str, str] = {
    ".docm": ".docx",
    ".xlsm": ".xlsx",
    ".tsv": ".csv",
    ".text": ".txt",
}
# Định dạng cũ (nhị phân) không đọc được bằng các thư viện trên: báo lỗi rõ ràng
# kèm hướng dẫn lưu lại, thay vì để thư viện ném ra lỗi khó hiểu.
LEGACY_EXTENSIONS: dict[str, str] = {".doc": ".docx", ".xls": ".xlsx", ".xlsb": ".xlsx"}
SUPPORTED_EXTENSIONS: frozenset[str] = frozenset(FORMAT_LABELS)
# Dùng khi tên file không có đuôi (một số client chỉ gửi content-type).
CONTENT_TYPE_EXTENSIONS: dict[str, str] = {
    "application/pdf": ".pdf",
    "application/x-pdf": ".pdf",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.ms-word.document.macroenabled.12": ".docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsx",
    "text/plain": ".txt",
    "text/csv": ".csv",
    "application/csv": ".csv",
    "text/tab-separated-values": ".csv",
}
# Thứ tự thử giải mã file text/csv: UTF-8 (kèm BOM) trước, rồi bảng mã tiếng Việt
# và Windows; latin-1 luôn giải mã được nên là chốt cuối.
TEXT_ENCODINGS: tuple[str, ...] = ("utf-8-sig", "cp1258", "cp1252", "latin-1")


class DocumentFormatError(ValueError):
    """Nội dung file không khớp định dạng khai báo (lỗi phía client -> HTTP 400)."""


class DependencyMissingError(RuntimeError):
    """Server thiếu thư viện đọc định dạng (lỗi cấu hình -> HTTP 503)."""


def _upload_suffix(filename: str, content_type: str = "") -> str:
    """Đuôi chuẩn (``.pdf``/``.docx``/``.xlsx``/``.txt``/``.csv``) của file upload.

    Ưu tiên đuôi trong tên file; nếu tên không có đuôi thì dùng content-type.
    Raise ``HTTPException`` 415 kèm danh sách định dạng hỗ trợ khi không đọc được,
    và hướng dẫn cụ thể với các định dạng cũ (``.doc``/``.xls``).
    """
    suffix = Path(filename or "").suffix.lower()
    canonical = EXTENSION_ALIASES.get(suffix, suffix)

    if canonical in SUPPORTED_EXTENSIONS:
        return canonical

    if suffix in LEGACY_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=(
                f"Định dạng {suffix} (bản cũ) chưa được hỗ trợ. "
                f"Hãy mở file bằng Word/Excel và lưu lại thành {LEGACY_EXTENSIONS[suffix]}."
            ),
        )

    media = (content_type or "").split(";")[0].strip().lower()
    media_suffix = CONTENT_TYPE_EXTENSIONS.get(media)
    if media_suffix:
        return EXTENSION_ALIASES.get(media_suffix, media_suffix)

    supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
    received = suffix or media or "không rõ"
    raise HTTPException(
        status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
        detail=(
            f"Định dạng file không được hỗ trợ ({received}). "
            f"Chỉ chấp nhận: {supported}."
        ),
    )


def _looks_like_pdf(data: bytes) -> bool:
    """A real PDF starts with the %PDF- signature (readers tolerate leading junk)."""
    return b"%PDF-" in data[:1024]


async def _read_upload(file: UploadFile, limit: int) -> bytes:
    """Read the upload in chunks, aborting as soon as the size limit is exceeded."""
    chunks: list[bytes] = []
    size = 0
    while chunk := await file.read(1024 * 1024):
        size += len(chunk)
        if size > limit:
            raise HTTPException(
                status_code=HTTP_413_CONTENT_TOO_LARGE,
                detail=(
                    f"File {file.filename or 'tải lên'} vượt quá giới hạn "
                    f"{limit // (1024 * 1024)} MB."
                ),
            )
        chunks.append(chunk)
    return b"".join(chunks)


async def _close_upload(file: UploadFile) -> None:
    """Close the spooled file, ignoring errors so one upload cannot break the batch."""
    try:
        await file.close()
    except Exception:  # noqa: BLE001 - closing is best effort
        logger.debug("Không đóng được upload %s", file.filename, exc_info=True)


def _open_pdf_reader(pdf_bytes: bytes) -> PdfReader:
    """Mở PDF (giải mã nếu cần) và trả về reader; lỗi -> ``ValueError`` rõ ràng."""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
    except PdfReadError as exc:
        raise ValueError(f"Không đọc được file PDF: {exc}") from exc
    except Exception as exc:  # noqa: BLE001 - pypdf can raise many low level errors
        raise ValueError(f"File PDF bị hỏng hoặc không hợp lệ: {exc}") from exc

    if reader.is_encrypted:
        # Try the common "empty password" case, otherwise give up clearly.
        try:
            if not reader.decrypt(""):
                raise ValueError(
                    "File PDF được bảo vệ bằng mật khẩu, không thể trích xuất nội dung."
                )
        except ValueError:
            raise
        except Exception as exc:  # noqa: BLE001
            raise ValueError(
                "File PDF được bảo vệ bằng mật khẩu, không thể trích xuất nội dung."
            ) from exc

    if not reader.pages:
        raise ValueError("File PDF không có trang nào.")
    return reader


def extract_pdf_pages(pdf_bytes: bytes) -> list[str]:
    """Text của TỪNG trang PDF theo đúng thứ tự (trang không có text -> chuỗi rỗng).

    ``extract_pdf_text`` dùng hàm này để giữ nguyên hành vi cũ, còn
    ``/process-book`` cần ranh giới trang để cắt sách thành từng chương.
    """
    reader = _open_pdf_reader(pdf_bytes)

    pages: list[str] = []
    for index, page in enumerate(reader.pages, start=1):
        try:
            page_text = page.extract_text() or ""
        except Exception as exc:  # noqa: BLE001 - a single broken page should not fail the run
            logger.warning("Không trích xuất được trang %s: %s", index, exc)
            page_text = ""
        pages.append(page_text)
    return pages


def extract_pdf_text(pdf_bytes: bytes) -> tuple[str, int]:
    """Extract text from every page with pypdf.

    Returns ``(text, page_count)``. Raises ``ValueError`` for unusable PDFs so the
    route can translate it into a 4xx response.
    """
    pages = extract_pdf_pages(pdf_bytes)
    parts = [
        f"--- Trang {index} ---\n{page_text.strip()}"
        for index, page_text in enumerate(pages, start=1)
        if page_text.strip()
    ]
    return "\n\n".join(parts), len(pages)


def _require_docx() -> None:
    """Báo lỗi cấu hình rõ ràng khi thiếu python-docx."""
    if DocxDocument is None:
        raise DependencyMissingError(
            "Server chưa cài python-docx nên không đọc được file Word. "
            "Chạy: pip install python-docx (hoặc pip install -r requirements.txt)."
        )


def _require_openpyxl() -> None:
    """Báo lỗi cấu hình rõ ràng khi thiếu openpyxl."""
    if load_xlsx_workbook is None:
        raise DependencyMissingError(
            "Server chưa cài openpyxl nên không đọc được file Excel. "
            "Chạy: pip install openpyxl (hoặc pip install -r requirements.txt)."
        )


def _zip_members(data: bytes) -> set[str]:
    """Tên các thành phần bên trong file zip; rỗng nếu không phải zip hợp lệ."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            return set(archive.namelist())
    except (zipfile.BadZipFile, OSError):
        return set()


def _guard_zip_expansion(data: bytes, limit: int) -> None:
    """Chặn "zip bomb": file nén nhỏ nhưng giải nén ra hàng trăm MB."""
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            total = sum(info.file_size for info in archive.infolist())
    except (zipfile.BadZipFile, OSError) as exc:
        raise DocumentFormatError(
            "File không phải kho nén hợp lệ (docx/xlsx bị hỏng?)."
        ) from exc
    if total > limit * MAX_ZIP_EXPANSION_RATIO:
        raise ValueError(
            f"Nội dung giải nén của file quá lớn ({total // (1024 * 1024)} MB), "
            "từ chối xử lý để bảo vệ server."
        )


def decode_text_bytes(data: bytes) -> str:
    """Giải mã file text/csv: UTF-8 (có/không BOM), UTF-16 có BOM, rồi bảng mã Việt."""
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        try:
            return data.decode("utf-16")
        except UnicodeDecodeError:
            pass
    for encoding in TEXT_ENCODINGS:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def estimate_pages(text: str) -> int:
    """Số trang ước lượng cho định dạng không có khái niệm trang (docx/txt/csv)."""
    return max(1, math.ceil(len(text) / max(1, CHARS_PER_PAGE_ESTIMATE)))


def _join_row(cells: Iterable[str]) -> str:
    """Ghép các ô của một dòng thành text một dòng, bỏ ô rỗng ở cuối."""
    values = list(cells)
    while values and not values[-1]:
        values.pop()
    return " | ".join(values)


def _cell_text(value: Any) -> str:
    """Một ô Excel -> text một dòng (None thành chuỗi rỗng)."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, dt.datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, (dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return " ".join(str(value).split())


def extract_text_file(data: bytes) -> tuple[str, int]:
    """Đọc ``.txt``: tự giải mã bảng mã và từ chối nội dung nhị phân.

    Returns ``(text, page_count)`` như ``extract_pdf_text``.
    """
    text = decode_text_bytes(data)
    if "\x00" in text:
        raise DocumentFormatError(
            "File .txt chứa dữ liệu nhị phân nên không phải văn bản thuần."
        )
    normalized = text.replace("\r\n", "\n").replace("\r", "\n").strip()
    return normalized, estimate_pages(normalized)


def extract_csv_text(data: bytes) -> tuple[str, int]:
    """Đọc ``.csv``/``.tsv`` bằng ``csv`` của stdlib.

    Tự dò dấu phân tách (`,`, `;`, tab, `|`) để chịu được CSV xuất từ Excel
    tiếng Việt (thường dùng `;`), rồi render mỗi dòng thành ``ô | ô | ô`` để
    model nhìn thấy cấu trúc bảng.
    """
    text = decode_text_bytes(data)
    if "\x00" in text:
        raise DocumentFormatError(
            "File CSV chứa dữ liệu nhị phân nên không phải văn bản thuần."
        )
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    sample = normalized[:8192]

    delimiter = ","
    try:
        delimiter = csv.Sniffer().sniff(sample, delimiters=",;\t|").delimiter
    except csv.Error:
        # Sniffer thất bại (file 1 cột / quá ít dòng): chọn dấu xuất hiện nhiều
        # nhất trên dòng đầu, mặc định dấu phẩy.
        first_line = sample.split("\n", 1)[0]
        counts = {
            separator: first_line.count(separator) for separator in (",", ";", "\t", "|")
        }
        best = max(counts, key=lambda separator: counts[separator])
        delimiter = best if counts[best] else ","

    rows: list[str] = []
    try:
        for record in csv.reader(io.StringIO(normalized), delimiter=delimiter):
            row = _join_row(" ".join(cell.split()) for cell in record)
            if row.strip(" |"):
                rows.append(row)
    except csv.Error as exc:
        raise ValueError(f"Không đọc được file CSV: {exc}") from exc

    rendered = "\n".join(rows)
    return rendered, estimate_pages(rendered)


def extract_docx_text(data: bytes) -> tuple[str, int]:
    """Đọc ``.docx``/``.docm`` bằng python-docx: đoạn văn trước, rồi từng bảng.

    Bảng được render thành các dòng ``ô | ô`` (kèm tiêu đề ``[Bảng n]``) để model
    vẫn thấy cấu trúc khi sinh câu hỏi dạng bảng.
    """
    _require_docx()
    assert DocxDocument is not None  # giúp type checker, _require_docx đã kiểm tra

    try:
        document = DocxDocument(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001 - python-docx ném nhiều loại lỗi
        raise ValueError(f"Không đọc được file Word (.docx): {exc}") from exc

    blocks: list[str] = []
    for paragraph in document.paragraphs:
        paragraph_text = paragraph.text.strip()
        if paragraph_text:
            blocks.append(paragraph_text)

    for index, table in enumerate(document.tables, start=1):
        rows: list[str] = []
        for row in table.rows:
            rendered = _join_row(" ".join(cell.text.split()) for cell in row.cells)
            if rendered.strip(" |"):
                rows.append(rendered)
        if rows:
            blocks.append(f"[Bảng {index}]\n" + "\n".join(rows))

    text = "\n\n".join(blocks)
    return text, estimate_pages(text)


def extract_xlsx_text(data: bytes) -> tuple[str, int]:
    """Đọc ``.xlsx``/``.xlsm`` bằng openpyxl, mỗi worksheet thành một khối text.

    Returns ``(text, sheet_count)``: với Excel, "số trang" chính là số sheet.
    """
    _require_openpyxl()
    assert load_xlsx_workbook is not None  # _require_openpyxl đã kiểm tra

    try:
        workbook = load_xlsx_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001 - openpyxl ném nhiều loại lỗi
        raise ValueError(f"Không đọc được file Excel (.xlsx): {exc}") from exc

    try:
        blocks: list[str] = []
        sheet_count = 0
        for sheet in workbook.worksheets:
            sheet_count += 1
            rows: list[str] = []
            for row in sheet.iter_rows(values_only=True):
                rendered = _join_row(_cell_text(value) for value in row)
                if rendered.strip(" |"):
                    rows.append(rendered)
            header = f"### Sheet: {sheet.title}"
            blocks.append(header if not rows else f"{header}\n" + "\n".join(rows))
    finally:
        workbook.close()

    text = "\n\n".join(blocks)
    return text, max(1, sheet_count)


def extract_text(filename: str, data: bytes, content_type: str = "") -> tuple[str, int]:
    """Chuyển nội dung MỘT file (đã đọc vào bộ nhớ) thành plain text.

    Tự nhận diện định dạng từ đuôi file (``.pdf``, ``.docx``, ``.xlsx``, ``.txt``,
    ``.csv`` + alias) hoặc từ ``content_type`` khi tên file không có đuôi, rồi
    gọi đúng thư viện đọc. Trả về ``(text, page_count)``.

    Lỗi được phân loại để route trả về thông báo rõ ràng thay vì làm sập server:
        - ``HTTPException`` 415: định dạng không được hỗ trợ
        - ``DocumentFormatError``: nội dung không khớp định dạng (400)
        - ``DependencyMissingError``: server thiếu thư viện (503)
        - ``ValueError``: file hỏng / không đọc được (422)
    """
    suffix = _upload_suffix(filename, content_type)
    label = FORMAT_LABELS[suffix]

    try:
        if suffix == ".pdf":
            if not _looks_like_pdf(data):
                raise DocumentFormatError("Nội dung file không phải định dạng PDF hợp lệ.")
            return extract_pdf_text(data)

        if suffix == ".docx":
            _guard_zip_expansion(data, MAX_UPLOAD_BYTES)
            if "word/document.xml" not in _zip_members(data):
                raise DocumentFormatError("Nội dung file không phải Word (.docx) hợp lệ.")
            return extract_docx_text(data)

        if suffix == ".xlsx":
            _guard_zip_expansion(data, MAX_UPLOAD_BYTES)
            if "xl/workbook.xml" not in _zip_members(data):
                raise DocumentFormatError("Nội dung file không phải Excel (.xlsx) hợp lệ.")
            return extract_xlsx_text(data)

        if suffix == ".csv":
            return extract_csv_text(data)

        return extract_text_file(data)
    except (DocumentFormatError, DependencyMissingError, HTTPException):
        raise
    except ValueError:
        raise
    except Exception as exc:  # noqa: BLE001 - mọi lỗi thư viện thành lỗi 4xx rõ ràng
        logger.warning(
            "Không trích xuất được %s (%s): %s", filename, label, exc, exc_info=True
        )
        raise ValueError(f"Không trích xuất được nội dung từ file {label}: {exc}") from exc


def _empty_text_message(suffix: str) -> str:
    """Thông báo riêng cho từng định dạng khi trích xuất ra text rỗng."""
    if suffix == ".pdf":
        return (
            "Không trích xuất được text từ PDF. File có thể là bản scan/ảnh "
            "và cần OCR trước khi xử lý."
        )
    if suffix == ".docx":
        return (
            "File Word không có nội dung văn bản (có thể chỉ chứa ảnh/shape). "
            "Hãy kiểm tra lại file rồi gửi lại."
        )
    if suffix == ".xlsx":
        return "File Excel không có dữ liệu trong ô nào."
    return "File không có nội dung văn bản để xử lý."


def _repair_json_response(text: str) -> Any | None:
    """Best-effort repair of truncated / malformed JSON via ``json-repair``.

    ``json-repair`` closes missing ``}``/``]``/quotes and drops markdown fences or
    prose around the payload, which is exactly the shape of a response that was
    cut off by the token limit. Returns the repaired object/array, or ``None``
    when the library is missing or the text is not usable JSON (plain prose …).
    """
    if _json_repair is None:
        return None

    # Repair the whole answer and, when the model prefixed prose, the payload
    # starting at the first "{" or "[" (that also covers a truncated tail).
    candidates = [text]
    first_positions = [pos for pos in (text.find("{"), text.find("[")) if pos != -1]
    if first_positions and min(first_positions) > 0:
        candidates.append(text[min(first_positions) :])

    for candidate in candidates:
        try:
            repaired = _json_repair(candidate, return_objects=True)
        except Exception:  # noqa: BLE001 - repair is best-effort, never fatal
            logger.warning("json-repair không sửa được phản hồi của model.", exc_info=True)
            continue
        # "" / None / scalar means json-repair found nothing worth keeping.
        if isinstance(repaired, (dict, list)) and repaired:
            logger.warning(
                "Đã sửa JSON hỏng/bị cắt cụt bằng json-repair (%s -> %s ký tự).",
                len(candidate),
                len(json.dumps(repaired, ensure_ascii=False)),
            )
            return repaired
    return None


def parse_json_response(raw: str) -> Any:
    """Parse the model output as JSON, tolerating fences, prose and truncation.

    Thứ tự thử: JSON thuần -> cắt đoạn ``{…}`` đầu tiên -> ``json-repair`` (tự
    đóng ngoặc/chuỗi bị cắt cụt). Raises ``json.JSONDecodeError`` when no JSON
    value can be recovered.
    """
    text = raw.strip()

    fenced = re.search(r"```(?:json)?\s*(.+?)```", text, flags=re.DOTALL | re.IGNORECASE)
    if fenced:
        text = fenced.group(1).strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError as first_error:
        last_error = first_error

    # Slice from the first "{" to the last "}" and retry (prose around the JSON).
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        try:
            return json.loads(text[start : end + 1])
        except json.JSONDecodeError as exc:
            last_error = exc

    repaired = _repair_json_response(text)
    if repaired is not None:
        return repaired

    raise last_error


# --------------------------------------------------------------------------------------
# Question / curriculum normalisation (shared by /process-document and /generate-exam)
# --------------------------------------------------------------------------------------
@dataclass(slots=True)
class ExtractedDocument:
    """One successfully extracted document of the subject (kept in memory for the corpus)."""

    position: int
    filename: str
    text: str
    pages: int
    characters: int
    truncated: bool
    # Định dạng gốc: "pdf", "docx", "xlsx", "txt", "csv" (mặc định pdf để tương thích).
    format: str = "pdf"


class QuestionShortfallError(RuntimeError):
    """The model returned fewer questions than requested, even after top-ups."""

    def __init__(self, requested: int, produced: int) -> None:
        super().__init__(
            f"Model chỉ tạo được {produced}/{requested} câu hỏi kể cả sau khi yêu cầu bổ sung."
        )
        self.requested = requested
        self.produced = produced


def _as_text(value: Any) -> str:
    """Coerce a model value to a stripped string (empty for None/containers)."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (int, float)):
        return str(value)
    return ""


def _as_text_list(value: Any) -> list[str]:
    """Coerce a string / newline separated string / list into a list of strings."""
    if value is None:
        return []
    if isinstance(value, str):
        parts = re.split(r"[\n;]", value)
        return [part.strip() for part in parts if part.strip()]
    if isinstance(value, (list, tuple, set)):
        return [text for item in value if (text := _as_text(item))]
    text = _as_text(value)
    return [text] if text else []


def normalize_key_concepts(value: Any) -> list[dict[str, Any]]:
    """Normalise ``key_concepts`` to ``[{name, explanation, sources}]``."""
    items = value if isinstance(value, (list, tuple)) else [value] if value else []
    concepts: list[dict[str, Any]] = []
    for item in items:
        if isinstance(item, str):
            name = item.strip()
            if name:
                concepts.append({"name": name, "explanation": "", "sources": []})
            continue
        if not isinstance(item, dict):
            continue
        name = _as_text(item.get("name") or item.get("concept") or item.get("title"))
        explanation = _as_text(item.get("explanation") or item.get("description"))
        if not name and not explanation:
            continue
        concepts.append(
            {
                "name": name or explanation[:80],
                "explanation": explanation,
                "sources": _as_text_list(item.get("sources") or item.get("source")),
            }
        )
    return concepts


def normalize_formulas(value: Any) -> list[dict[str, Any]]:
    """Normalise ``formulas`` to ``[{formula, description, symbols, sources}]``."""
    items = value if isinstance(value, (list, tuple)) else [value] if value else []
    formulas: list[dict[str, Any]] = []
    for item in items:
        if isinstance(item, str):
            if item.strip():
                formulas.append(
                    {"formula": item.strip(), "description": "", "symbols": [], "sources": []}
                )
            continue
        if not isinstance(item, dict):
            continue
        formula = _as_text(item.get("formula") or item.get("expression") or item.get("name"))
        description = _as_text(item.get("description") or item.get("meaning"))
        if not formula and not description:
            continue
        raw_symbols = item.get("symbols")
        if isinstance(raw_symbols, dict):
            raw_symbols = [
                {"symbol": key, "meaning": meaning} for key, meaning in raw_symbols.items()
            ]
        symbols: list[dict[str, str]] = []
        for symbol in raw_symbols if isinstance(raw_symbols, (list, tuple)) else []:
            if isinstance(symbol, str):
                if symbol.strip():
                    symbols.append({"symbol": symbol.strip(), "meaning": ""})
            elif isinstance(symbol, dict):
                key = _as_text(symbol.get("symbol") or symbol.get("name"))
                meaning = _as_text(symbol.get("meaning") or symbol.get("description"))
                if key or meaning:
                    symbols.append({"symbol": key, "meaning": meaning})
        formulas.append(
            {
                "formula": formula,
                "description": description,
                "symbols": symbols,
                "sources": _as_text_list(item.get("sources") or item.get("source")),
            }
        )
    return formulas


def normalize_examples(value: Any) -> list[dict[str, Any]]:
    """Normalise ``examples`` to ``[{title, problem, solution, sources}]``."""
    items = value if isinstance(value, (list, tuple)) else [value] if value else []
    examples: list[dict[str, Any]] = []
    for index, item in enumerate(items, start=1):
        if isinstance(item, str):
            if item.strip():
                examples.append(
                    {
                        "title": f"Ví dụ {index}",
                        "problem": item.strip(),
                        "solution": "",
                        "sources": [],
                    }
                )
            continue
        if not isinstance(item, dict):
            continue
        problem = _as_text(item.get("problem") or item.get("question") or item.get("content"))
        solution = _as_text(
            item.get("solution") or item.get("explanation") or item.get("answer")
        )
        if not problem and not solution:
            continue
        examples.append(
            {
                "title": _as_text(item.get("title")) or f"Ví dụ {index}",
                "problem": problem,
                "solution": solution,
                "sources": _as_text_list(item.get("sources") or item.get("source")),
            }
        )
    return examples


def normalize_cross_document_links(value: Any) -> list[dict[str, Any]]:
    """Normalise ``cross_document_links`` to ``[{documents, relation}]``."""
    items = value if isinstance(value, (list, tuple)) else [value] if value else []
    links: list[dict[str, Any]] = []
    for item in items:
        if isinstance(item, str):
            if item.strip():
                links.append({"documents": [], "relation": item.strip()})
            continue
        if not isinstance(item, dict):
            continue
        documents = _as_text_list(item.get("documents") or item.get("sources"))
        relation = _as_text(
            item.get("relation") or item.get("description") or item.get("note")
        )
        if relation or documents:
            links.append({"documents": documents, "relation": relation})
    return links


# Khoá gốc (chưa có hậu tố ngôn ngữ) mà model có thể dùng cho mảng khái niệm.
_CONCEPT_KEYS: frozenset[str] = frozenset(
    {"concepts", "concept", "key_concepts", "keyconcepts", "begriffe"}
)


def _concept_language(normalized_key: str) -> str | None:
    """Ngôn ngữ của khoá khái niệm (``concepts_de`` / ``keyConceptsVi`` / ...)."""
    for code in BILINGUAL_LANGUAGES:
        for suffix in (f"_{code}", code):
            stem = normalized_key[: -len(suffix)]
            if normalized_key.endswith(suffix) and stem in _CONCEPT_KEYS:
                return code
    return None


def _project_concept_item(item: dict[str, Any], code: str) -> dict[str, Any]:
    """Chọn bản ``de``/``vi`` bên trong MỘT khái niệm có trường song ngữ."""
    projected: dict[str, Any] = {}
    for key, value in item.items():
        if isinstance(value, dict) and any(language_code(inner) for inner in value):
            text = localized_text(value, lang=code)[code]
            if text:
                projected[key] = text
            continue
        projected[key] = value
    return projected


def _has_localized_items(value: Any) -> bool:
    """True khi mảng khái niệm có trường con dạng ``{"de": ..., "vi": ...}``."""
    items = value if isinstance(value, (list, tuple)) else [value]
    return any(
        isinstance(item, dict)
        and any(
            isinstance(field, dict) and any(language_code(key) for key in field)
            for field in item.values()
        )
        for item in items
    )


def _project_concepts(value: Any, code: str) -> list[Any]:
    """Danh sách khái niệm của một ngôn ngữ, kể cả khi giá trị là object song ngữ."""
    if isinstance(value, dict):
        if any(language_code(key) for key in value):
            for key, item in value.items():
                if language_code(key) == code:
                    return _project_concepts(item, code)
            return []
        # Một khái niệm đơn lẻ ({"name": ..., "explanation": ...}).
        return [_project_concept_item(value, code)]
    items = list(value) if isinstance(value, (list, tuple)) else [value] if value else []
    return [
        _project_concept_item(item, code) if isinstance(item, dict) else item
        for item in items
    ]


def normalize_bilingual_summary(
    teaching: dict[str, Any], *, lang: str = DEFAULT_OUTPUT_LANGUAGE, fallback: str = ""
) -> dict[str, str]:
    """Bản tóm tắt song ngữ ``{"de": ..., "vi": ...}``.

    Chấp nhận ``summary_de``/``summary_vi``, object ``summary: {"de", "vi"}`` và
    cả ``summary``/``overview`` dạng chuỗi của hợp đồng cũ (xếp vào bản của ``lang``).
    """
    summary = _new_bilingual_text()
    for key, value in teaching.items():
        normalized_key = _section_key(key)
        for code in BILINGUAL_LANGUAGES:
            if normalized_key in {f"summary_{code}", f"summary{code}"} and not summary[code]:
                summary[code] = _section_text(value)

    if not any(summary.values()):
        summary = localized_text(teaching.get("summary") or teaching.get("overview"), lang=lang)

    primary = resolve_language(lang)
    if not summary[primary] and _as_text(fallback):
        summary[primary] = _as_text(fallback)
    return summary


def normalize_bilingual_concepts(
    teaching: dict[str, Any], *, lang: str = DEFAULT_OUTPUT_LANGUAGE
) -> dict[str, list[dict[str, Any]]]:
    """Hai mảng khái niệm ``{"de": [...], "vi": [...]}``.

    Chấp nhận ``concepts_de``/``concepts_vi``, ``key_concepts_de``/``key_concepts_vi``,
    object ``concepts: {"de": [...], "vi": [...]}`` và danh sách một ngôn ngữ
    (``key_concepts`` của hợp đồng cũ) — danh sách một ngôn ngữ được xếp vào bản
    của ``lang`` để client cũ vẫn đọc được.
    """
    result = {code: [] for code in BILINGUAL_LANGUAGES}
    generic: list[Any] = []
    primary = resolve_language(lang)

    for key, value in teaching.items():
        normalized_key = _section_key(key)
        matched_lang = _concept_language(normalized_key)
        if matched_lang:
            result[matched_lang] = normalize_key_concepts(
                _project_concepts(value, matched_lang)
            )
            continue
        if normalized_key in _CONCEPT_KEYS:
            generic.append(value)

    for value in generic:
        if isinstance(value, dict) and any(language_code(key) for key in value):
            for code in BILINGUAL_LANGUAGES:
                if not result[code]:
                    result[code] = normalize_key_concepts(_project_concepts(value, code))
            continue
        if _has_localized_items(value):
            # Một mảng nhưng mỗi khái niệm tự chứa bản de/vi -> tách cho cả hai.
            for code in BILINGUAL_LANGUAGES:
                if not result[code]:
                    result[code] = normalize_key_concepts(_project_concepts(value, code))
            continue
        if not result[primary]:
            result[primary] = normalize_key_concepts(_project_concepts(value, primary))

    return result



def normalize_teaching(
    value: Any, *, fallback_summary: str = "", lang: str = DEFAULT_OUTPUT_LANGUAGE
) -> dict[str, Any]:
    """Guarantee the full ``teaching`` shape even if the model omits a field.

    ``summary_de``/``summary_vi`` và ``concepts_de``/``concepts_vi`` là hợp đồng
    chính; ``summary``/``key_concepts`` (không hậu tố) được giữ lại như bản theo
    ``lang`` để client cũ không phải sửa ngay.
    """
    teaching = value if isinstance(value, dict) else {}
    summary = normalize_bilingual_summary(
        teaching, lang=lang, fallback=fallback_summary
    )
    concepts = normalize_bilingual_concepts(teaching, lang=lang)
    primary = resolve_language(lang)
    secondary = next((code for code in BILINGUAL_LANGUAGES if code != primary), primary)
    return {
        "summary": summary[primary] or summary[secondary],
        "summary_de": summary["de"],
        "summary_vi": summary["vi"],
        "key_concepts": concepts[primary] or concepts[secondary],
        "concepts_de": concepts["de"],
        "concepts_vi": concepts["vi"],
        "formulas": normalize_formulas(teaching.get("formulas")),
        "examples": normalize_examples(teaching.get("examples")),
        "cross_document_links": normalize_cross_document_links(
            teaching.get("cross_document_links") or teaching.get("crossDocumentLinks")
        ),
    }


_OPTION_LETTERS = "ABCDEFGH"


def normalize_options(value: Any) -> dict[str, str] | None:
    """Normalise the ``options`` field (dict or list) to ``{"A": "...", ...}``."""
    if isinstance(value, dict):
        options: dict[str, str] = {}
        for index, (key, label) in enumerate(value.items()):
            letter = _as_text(key).upper()
            if not letter:
                letter = _OPTION_LETTERS[min(index, len(_OPTION_LETTERS) - 1)]
            text = _as_text(label)
            if text:
                options[letter] = text
        return options or None
    if isinstance(value, (list, tuple)):
        options = {}
        for index, label in enumerate(value):
            text = _as_text(label)
            if text and index < len(_OPTION_LETTERS):
                options[_OPTION_LETTERS[index]] = text
        return options or None
    return None


def infer_question_type(item: dict[str, Any]) -> str:
    """Best-effort type for a question the model returned without a valid ``type``."""
    explicit = normalize_question_type(item.get("type"))
    if explicit:
        return explicit
    if isinstance(item.get("table_data"), dict):
        return "table"
    options = item.get("options")
    answer = item.get("answer")
    if isinstance(options, (dict, list, tuple)) and options:
        if isinstance(answer, (list, tuple)) and len(answer) > 1:
            return "multiple_response"
        return "multiple_choice"
    if isinstance(answer, bool):
        return "true_false"
    if isinstance(answer, str) and answer.strip().lower() in {
        "true",
        "false",
        "đúng",
        "sai",
    }:
        return "true_false"
    if isinstance(item.get("given"), (list, tuple)) or _as_text(item.get("formula")):
        return "calculation"
    # A free-text answer fits the fill-in-the-blank contract best.
    return "fill_in_the_blank"


def normalize_question(item: Any) -> dict[str, Any] | None:
    """Coerce one raw model question to the documented contract, or ``None``."""
    if not isinstance(item, dict):
        return None
    question_text = _as_text(item.get("question") or item.get("prompt") or item.get("text"))
    if not question_text:
        return None

    qtype = infer_question_type(item)
    answer = item.get("answer") if item.get("answer") is not None else ""
    normalized: dict[str, Any] = {
        "id": 0,  # reassigned from the question plan order
        "type": qtype,
        "question": question_text,
        "answer": answer,
        "explanation": _as_text(item.get("explanation") or item.get("solution")),
        "difficulty": _as_text(item.get("difficulty")) or "Trung bình",
        "sources": _as_text_list(item.get("sources") or item.get("source")),
    }

    if qtype in {"multiple_choice", "multiple_response"}:
        options = normalize_options(item.get("options"))
        if not options:
            # Không có lựa chọn thì web không thể chấm tự động -> bỏ câu này.
            return None
        normalized["options"] = options
        if isinstance(answer, str):
            letters = [part.upper() for part in re.split(r"[,\s;/]+", answer) if part.strip()]
        elif isinstance(answer, (list, tuple)):
            letters = [str(part).strip().upper() for part in answer if str(part).strip()]
        else:
            letters = []
        normalized["answer"] = (
            letters[0] if qtype == "multiple_choice" and letters else letters
        )
    elif qtype == "calculation":
        normalized["given"] = _as_text_list(item.get("given") or item.get("data"))
        normalized["formula"] = _as_text(item.get("formula"))
        normalized["answer"] = _as_text(answer) or answer
    elif qtype == "fill_in_the_blank":
        if isinstance(answer, str):
            answers = [answer.strip()] if answer.strip() else []
        elif isinstance(answer, (list, tuple)):
            answers = [str(part).strip() for part in answer if str(part).strip()]
        else:
            answers = []
        normalized["answer"] = answers
    elif qtype == "true_false":
        if isinstance(answer, bool):
            normalized["answer"] = "true" if answer else "false"
        else:
            truthy = _as_text(answer).lower() in {"true", "đúng", "dung", "yes", "1"}
            normalized["answer"] = "true" if truthy else "false"
    elif qtype == "table":
        table_data = item.get("table_data") if isinstance(item.get("table_data"), dict) else {}
        headers = _as_text_list(table_data.get("headers"))
        rows = [
            [_as_text(cell) for cell in row]
            for row in (table_data.get("rows") or [])
            if isinstance(row, (list, tuple))
        ]
        if not headers or not rows:
            return None
        normalized["table_data"] = {"headers": headers, "rows": rows}
        cells: list[dict[str, Any]] = []
        for cell in answer if isinstance(answer, list) else []:
            if not isinstance(cell, dict):
                continue
            try:
                row_index = int(cell.get("row", 0))
                column_index = int(cell.get("column", 0))
            except (TypeError, ValueError):
                continue
            cells.append(
                {
                    "row": row_index,
                    "column": column_index,
                    "value": _as_text(cell.get("value")),
                }
            )
        normalized["answer"] = cells

    return normalized


def parse_questions(payload: Any) -> list[dict[str, Any]]:
    """Extract and normalise the ``questions`` array of a model payload."""
    if isinstance(payload, dict):
        items = payload.get("questions")
    elif isinstance(payload, (list, tuple)):
        items = payload
    else:
        items = None
    if not isinstance(items, (list, tuple)):
        return []
    parsed: list[dict[str, Any]] = []
    for item in items:
        normalized = normalize_question(item)
        if normalized is not None:
            parsed.append(normalized)
    return parsed


def question_stems(
    questions: Iterable[dict[str, Any]], limit: int = MAX_EXAM_EXISTING_QUESTIONS
) -> list[str]:
    """Question texts to send back as "do not repeat these" hints."""
    stems: list[str] = []
    for question in questions:
        stem = _as_text(question.get("question"))
        if stem and stem not in stems:
            stems.append(stem[:200])
    return stems[-limit:]


def _question_key(question: dict[str, Any]) -> str:
    """Fuzzy identity of a question, used to drop duplicates across model calls."""
    stem = re.sub(r"\s+", " ", _as_text(question.get("question")).lower()).strip()
    return re.sub(r"[^\w\s]", "", stem)[:160]


def dedupe_questions(questions: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    """Drop repeated/near repeated questions, keeping the first occurrence."""
    seen: set[str] = set()
    unique: list[dict[str, Any]] = []
    for question in questions:
        if not isinstance(question, dict):
            continue
        key = _question_key(question)
        if not key or key in seen:
            continue
        seen.add(key)
        unique.append(question)
    return unique


def align_questions_to_plan(
    questions: Iterable[dict[str, Any]], plan: Sequence[str]
) -> list[dict[str, Any]]:
    """Pick at most one question per plan slot, in plan order, renumbering ids.

    The model often ignores the requested mix, so every slot is matched with the
    first unused question of that type. Questions that fit no slot are dropped;
    ``complete_questions`` then asks the model for the missing ones.
    """
    pool = [question for question in questions if isinstance(question, dict)]
    used: set[int] = set()
    aligned: list[dict[str, Any]] = []
    for qtype in plan:
        for index, question in enumerate(pool):
            if index in used or infer_question_type(question) != qtype:
                continue
            used.add(index)
            aligned.append({**question, "type": qtype, "id": len(aligned) + 1})
            break
    return aligned


def question_type_deficits(
    aligned: Sequence[dict[str, Any]], plan: Sequence[str]
) -> list[tuple[str, int]]:
    """How many questions of each plan type are still missing (canonical order)."""
    missing = Counter(plan)
    for question in aligned:
        qtype = infer_question_type(question)
        if missing.get(qtype, 0) > 0:
            missing[qtype] -= 1
    return [
        (qtype, missing[qtype])
        for qtype in question_plan_counts(plan)
        if missing.get(qtype, 0) > 0
    ]


# --------------------------------------------------------------------------------------
# Routes
# --------------------------------------------------------------------------------------
@app.get("/health", tags=["system"])
async def health() -> dict[str, str]:
    """Liveness probe used by the frontend / container orchestrators."""
    return {
        "status": "ok",
        "model": DEEPSEEK_MODEL,
        "base_url": DEEPSEEK_BASE_URL,
    }


def request_completion(
    *,
    system_prompt: str,
    user_content: str,
    label: str,
    temperature: float = DEEPSEEK_TEMPERATURE,
    max_tokens: int = DEEPSEEK_MAX_TOKENS,
) -> Any:
    """Call DeepSeek with JSON output mode and translate SDK errors to HTTP errors.

    Shared by ``/process-document`` and ``/generate-exam`` so both routes behave
    identically when the key is missing, the network is down or the quota is
    exhausted. ``label`` only appears in the server log (file name / exam scope).
    """
    client = get_deepseek_client()

    try:
        return client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content},
            ],
            # DeepSeek supports OpenAI-compatible JSON output mode.
            response_format={"type": "json_object"},
            max_tokens=max_tokens,
            temperature=temperature,
            stream=False,
        )
    except APITimeoutError as exc:
        logger.exception("DeepSeek timeout cho %s", label)
        raise HTTPException(
            status_code=status.HTTP_504_GATEWAY_TIMEOUT,
            detail="DeepSeek phản hồi quá lâu, vui lòng thử lại.",
        ) from exc
    except APIConnectionError as exc:
        logger.exception("Không kết nối được DeepSeek cho %s", label)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Không kết nối được tới DeepSeek API.",
        ) from exc
    except APIStatusError as exc:
        logger.exception("DeepSeek trả về lỗi %s cho %s", exc.status_code, label)
        message = exc.message or str(exc)
        if exc.status_code in {401, 403}:
            detail = "DeepSeek API key không hợp lệ hoặc không có quyền truy cập."
        elif exc.status_code == 429:
            detail = "Đã vượt hạn mức gọi DeepSeek, vui lòng thử lại sau."
        else:
            detail = f"DeepSeek API trả về lỗi {exc.status_code}: {message}"
        raise HTTPException(status_code=status.HTTP_502_BAD_GATEWAY, detail=detail) from exc
    except OpenAIError as exc:
        logger.exception("Lỗi DeepSeek SDK cho %s", label)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=f"Lỗi khi gọi DeepSeek: {exc}",
        ) from exc


def usage_dict(completion: Any) -> dict[str, Any]:
    """Token usage of a completion, tolerant of missing fields in test doubles."""
    usage = getattr(completion, "usage", None)

    return {
        "prompt_tokens": getattr(usage, "prompt_tokens", None),
        "completion_tokens": getattr(usage, "completion_tokens", None),
        "total_tokens": getattr(usage, "total_tokens", None),
    }


def merge_usage(usages: Iterable[dict[str, Any]]) -> dict[str, int | None]:
    """Sum the token usage of every DeepSeek call made for one document/subject."""
    keys = ("prompt_tokens", "completion_tokens", "total_tokens")
    totals: dict[str, int | None] = {key: 0 for key in keys}
    saw_value = False
    for usage in usages:
        if not isinstance(usage, dict):
            continue
        for key in keys:
            value = usage.get(key)
            if isinstance(value, int) and not isinstance(value, bool):
                totals[key] = (totals[key] or 0) + value
                saw_value = True
    if not saw_value:
        return {key: None for key in keys}
    return totals


def completion_text(completion: Any, label: str) -> str:
    """Return the assistant message of a completion, failing loudly when empty."""
    choices = getattr(completion, "choices", None) or []
    choice = choices[0] if choices else None
    raw_content = (getattr(choice.message, "content", None) or "") if choice else ""
    if not raw_content.strip():
        logger.error("DeepSeek trả về nội dung rỗng cho %s", label)
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="DeepSeek trả về nội dung rỗng.",
        )
    return raw_content


class JsonPayloadError(RuntimeError):
    """Model vẫn trả JSON hỏng sau khi json-repair đã sửa và đã gửi lại yêu cầu."""

    def __init__(self, *, label: str, raw_content: str, model: str, attempts: int) -> None:
        super().__init__(f"{label}: không có JSON hợp lệ sau {attempts} lần gọi DeepSeek.")
        self.label = label
        self.raw_content = raw_content
        self.model = model
        self.attempts = attempts


# Gắn vào user message của lần gửi lại: nhắc model đóng JSON và không viết thêm chữ.
JSON_RETRY_REMINDER = (
    "\n\nQUAN TRỌNG: phản hồi lần trước của bạn không phải JSON hợp lệ (có thể đã bị "
    "cắt cụt). Hãy trả lại TOÀN BỘ kết quả trong MỘT object JSON hoàn chỉnh: đóng đủ "
    "ngoặc {} [] và chuỗi, không viết thêm chữ nào ngoài JSON."
)


def request_json_payload(
    *,
    system_prompt: str,
    user_content: str,
    label: str,
    temperature: float = DEEPSEEK_TEMPERATURE,
    max_tokens: int = DEEPSEEK_MAX_TOKENS,
    retry_attempts: int | None = None,
    accept: tuple[type, ...] = (dict,),
) -> tuple[Any, list[dict[str, Any]], str, str]:
    """Ask DeepSeek for JSON, repairing it first and re-asking when still broken.

    Returns ``(payload, usages, model, raw_content)``. The answer is parsed with
    ``parse_json_response`` (which runs ``json-repair``); if that still fails, the
    same request is sent again up to ``retry_attempts`` extra times (default
    ``JSON_PARSE_RETRY_ATTEMPTS`` = 2) with a reminder to emit complete JSON.
    ``accept`` is the tuple of JSON types the caller can use; anything else is
    treated as a broken payload and retried. Raises ``JsonPayloadError`` when no
    usable payload arrives, so every call site can decide its own error message.
    """
    extra_attempts = JSON_PARSE_RETRY_ATTEMPTS if retry_attempts is None else retry_attempts
    total_attempts = 1 + max(0, extra_attempts)
    usages: list[dict[str, Any]] = []
    raw_content = ""
    model = DEEPSEEK_MODEL
    last_error: Exception | None = None

    for attempt in range(1, total_attempts + 1):
        content = user_content
        if attempt > 1:
            content = f"{user_content}{JSON_RETRY_REMINDER}"
            logger.warning(
                "%s: JSON chưa hợp lệ, gửi lại yêu cầu cho DeepSeek (lần %s/%s).",
                label,
                attempt,
                total_attempts,
            )

        completion = request_completion(
            system_prompt=system_prompt,
            user_content=content,
            label=label if attempt == 1 else f"{label}:retry-{attempt - 1}",
            temperature=temperature,
            max_tokens=max_tokens,
        )
        usages.append(usage_dict(completion))
        model = completion.model or model
        raw_content = completion_text(completion, label)

        try:
            payload = parse_json_response(raw_content)
        except json.JSONDecodeError as exc:
            last_error = exc
            logger.warning(
                "%s: không parse được JSON ở lần %s/%s (json-repair đã thử).",
                label,
                attempt,
                total_attempts,
            )
            continue

        if not isinstance(payload, accept):
            last_error = ValueError(
                f"{label}: JSON là {type(payload).__name__}, cần {accept}."
            )
            logger.warning("%s: JSON sai kiểu (%s), thử lại.", label, type(payload).__name__)
            continue

        if attempt > 1:
            logger.info("%s: JSON hợp lệ sau %s lần gọi DeepSeek.", label, attempt)
        return payload, usages, model, raw_content

    raise JsonPayloadError(
        label=label,
        raw_content=raw_content,
        model=model,
        attempts=total_attempts,
    ) from last_error


def truncate_text(text: str, limit: int) -> str:
    """Cut ``text`` to ``limit`` characters (no-op for non-positive limits)."""
    if limit <= 0:
        return ""
    if len(text) <= limit:
        return text
    return text[:limit]


def _render_document_block(document: ExtractedDocument, excerpt: str, role: str) -> str:
    """One labelled, budget-limited document section of the corpus prompt."""
    header = f"### {role}: {document.filename} ({document.pages} trang)"
    if len(excerpt) < len(document.text):
        header += f" [rút gọn còn {len(excerpt)}/{len(document.text)} ký tự]"
    return f"{header}\n{excerpt.strip()}"


def document_sources(
    documents: Sequence[ExtractedDocument], focus_position: int | None = None
) -> list[dict[str, Any]]:
    """Metadata of the files a result was built from (echoed back to the client)."""
    return [
        {
            "filename": document.filename,
            "position": document.position,
            "format": document.format,
            "pages": document.pages,
            "characters_extracted": document.characters,
            "truncated": document.truncated,
            "role": "focus" if focus_position == document.position else "reference",
        }
        for document in documents
    ]


def build_single_document_corpus(
    document: ExtractedDocument, *, limit: int = MAX_TEXT_CHARS
) -> str:
    """Render ONLY ``document`` — used by ``mode=per_file`` (independent analysis).

    Không kèm trích đoạn của bất kỳ file nào khác, nên prompt gửi DeepSeek chỉ
    chứa nội dung của đúng tài liệu đang xử lý.
    """
    excerpt = truncate_text(document.text, min(limit, MAX_TEXT_CHARS))
    return _render_document_block(document, excerpt, "NỘI DUNG TÀI LIỆU DUY NHẤT")


def build_subject_corpus(
    documents: Sequence[ExtractedDocument],
    *,
    focus_position: int | None = None,
    total_limit: int = MAX_SUBJECT_CONTEXT_CHARS,
    reference_limit: int = MAX_CROSS_DOCUMENT_CHARS,
) -> str:
    """Render every document of the subject, giving the focus file the larger share.

    Chỉ ``mode=combined`` (opt-in) dùng hàm này để tổng hợp cả môn. ``mode=per_file``
    dùng ``build_single_document_corpus`` nên mỗi file được phân tích độc lập.
    Total size stays within ``total_limit`` characters.
    """
    if not documents:
        return ""

    blocks: list[str] = []
    focus = next(
        (document for document in documents if document.position == focus_position), None
    )
    if focus is not None:
        others = [document for document in documents if document.position != focus_position]
        focus_limit = max(MIN_DOCUMENT_EXCERPT_CHARS, total_limit - reference_limit)
        blocks.append(
            _render_document_block(
                focus, truncate_text(focus.text, focus_limit), "TÀI LIỆU CHÍNH"
            )
        )
        if others:
            share = max(MIN_DOCUMENT_EXCERPT_CHARS, reference_limit // len(others))
            for document in others:
                blocks.append(
                    _render_document_block(
                        document,
                        truncate_text(document.text, min(share, MAX_TEXT_CHARS)),
                        "TÀI LIỆU THAM KHẢO CÙNG MÔN",
                    )
                )
        return "\n\n".join(blocks)

    share = max(MIN_DOCUMENT_EXCERPT_CHARS, total_limit // len(documents))
    for document in documents:
        blocks.append(
            _render_document_block(
                document,
                truncate_text(document.text, min(share, MAX_TEXT_CHARS)),
                "TÀI LIỆU TRONG MÔN",
            )
        )
    return "\n\n".join(blocks)


def complete_questions(
    *,
    questions: Iterable[dict[str, Any]],
    plan: Sequence[str],
    system_prompt: str,
    corpus: str,
    label: str,
    context_label: str = "TOÀN BỘ NỘI DUNG TÀI LIỆU",
    subject: str = "",
    focus_filename: str | None = None,
    directive: str | None = None,
    temperature: float = DEEPSEEK_TEMPERATURE,
    max_tokens: int = DEEPSEEK_MAX_TOKENS,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    """Return exactly ``len(plan)`` questions, asking the model for the missing ones.

    The first model answer rarely matches the requested type mix, so the pool is
    aligned with the plan and up to ``QUESTION_TOPUP_ATTEMPTS`` extra calls request
    a precise per-type shortfall. Returns ``(questions, usages, topup_rounds)``;
    a remaining shortfall raises ``QuestionShortfallError`` when
    ``STRICT_QUESTION_COUNT`` is enabled, otherwise it is only logged.
    """
    total = len(plan)
    pool = dedupe_questions(questions)
    aligned = align_questions_to_plan(pool, plan)
    usages: list[dict[str, Any]] = []
    rounds = 0

    while len(aligned) < total and rounds < QUESTION_TOPUP_ATTEMPTS:
        rounds += 1
        missing = total - len(aligned)
        logger.info(
            "%s: thiếu %s/%s câu hỏi, gọi bổ sung lần %s/%s.",
            label,
            missing,
            total,
            rounds,
            QUESTION_TOPUP_ATTEMPTS,
        )
        try:
            completion = request_completion(
                system_prompt=system_prompt,
                user_content=build_topup_prompt(
                    shortfall=missing,
                    plan=plan,
                    generated=aligned,
                    corpus=corpus,
                    subject=subject,
                    focus_filename=focus_filename,
                    context_label=context_label,
                    directive=directive,
                ),
                label=f"{label}:topup-{rounds}",
                temperature=temperature,
                max_tokens=max_tokens,
            )
        except Exception:  # noqa: BLE001 - a failed top-up must surface as a shortfall
            logger.exception("%s: không gọi được DeepSeek để bổ sung câu hỏi.", label)
            break

        usages.append(usage_dict(completion))
        try:
            payload = parse_json_response(completion_text(completion, label))
        except (json.JSONDecodeError, HTTPException):
            logger.warning("%s: câu bổ sung không phải JSON hợp lệ, dừng bổ sung.", label)
            continue

        added = parse_questions(payload)
        if not added:
            logger.warning("%s: model không trả thêm câu hỏi nào.", label)
            continue

        pool = dedupe_questions([*pool, *added])
        aligned = align_questions_to_plan(pool, plan)

    if len(aligned) < total:
        message = (
            f"{label}: chỉ tạo được {len(aligned)}/{total} câu hỏi sau "
            f"{rounds} lần bổ sung."
        )
        if STRICT_QUESTION_COUNT:
            raise QuestionShortfallError(total, len(aligned))
        logger.warning(message)
    else:
        logger.info("%s: đủ %s câu hỏi sau %s lần bổ sung.", label, total, rounds)

    return aligned, usages, rounds


async def _extract_upload(file: UploadFile, position: int) -> ExtractedDocument:
    """Kiểm tra MỘT file (PDF/DOCX/XLSX/TXT/CSV) và trích xuất text vào bộ nhớ.

    Định dạng được nhận diện từ đuôi file (fallback: content-type) rồi chuyển
    thành plain text bằng thư viện tương ứng. Kết quả là nội dung của đúng file
    này: ở ``mode=per_file`` nó được gửi một mình cho DeepSeek (phân tích độc
    lập), còn ``mode=combined`` mới ghép các file của môn lại. Mọi lỗi được báo
    bằng ``HTTPException`` để route ghi nhận thành một phần tử lỗi (file này hỏng
    không làm hỏng cả batch).
    """
    filename = file.filename or f"document-{position}"
    content_type = file.content_type or ""
    # Nhận diện định dạng TRƯỚC khi đọc để báo 415 ngay cho định dạng lạ.
    suffix = _upload_suffix(filename, content_type)
    label = FORMAT_LABELS[suffix]

    try:
        data = await _read_upload(file, MAX_UPLOAD_BYTES)
    finally:
        await _close_upload(file)

    if not data:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"File {label} rỗng.",
        )

    try:
        text, page_count = extract_text(filename, data, content_type)
    except DocumentFormatError as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc)
        ) from exc
    except DependencyMissingError as exc:
        logger.error("Thiếu thư viện để đọc %s: %s", filename, exc)
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE, detail=str(exc)
        ) from exc
    except HTTPException:
        raise
    except ValueError as exc:
        raise HTTPException(
            status_code=HTTP_422_UNPROCESSABLE_CONTENT, detail=str(exc)
        ) from exc

    if not text.strip():
        raise HTTPException(
            status_code=HTTP_422_UNPROCESSABLE_CONTENT,
            detail=_empty_text_message(suffix),
        )

    truncated = len(text) > MAX_TEXT_CHARS
    document_text = text[:MAX_TEXT_CHARS]
    if truncated:
        logger.warning(
            "Tài liệu %s có %s ký tự, cắt còn %s ký tự trước khi gửi DeepSeek.",
            filename,
            len(text),
            MAX_TEXT_CHARS,
        )

    logger.info(
        "Đã trích xuất %s %s / %s ký tự từ tài liệu %s/%s.",
        page_count,
        "trang" if suffix == ".pdf" else "sheet" if suffix == ".xlsx" else "khối",
        len(document_text),
        position,
        filename,
    )

    return ExtractedDocument(
        position=position,
        filename=filename,
        text=document_text,
        pages=page_count,
        characters=len(text),
        truncated=truncated,
        format=suffix.lstrip("."),
    )


def _error_element(filename: str, status_code: int, error: str, **extra: Any) -> dict[str, Any]:
    """One ``success: false`` result element (a failed file must not kill the batch)."""
    return {
        "success": False,
        "filename": filename,
        "status_code": status_code,
        "error": error,
        **extra,
    }


def _generate_curriculum(
    *,
    documents: Sequence[ExtractedDocument],
    focus: ExtractedDocument | None,
    subject: str,
    lang: str,
    plan: Sequence[str],
    mode: str,
) -> dict[str, Any]:
    """Gọi DeepSeek (kèm bổ sung nếu thiếu câu) và dựng một phần tử kết quả.

    ``mode=per_file`` (mặc định) xử lý ĐỘC LẬP: ``focus`` là file duy nhất và chỉ
    nội dung của file đó được gửi cho DeepSeek (``documents`` cũng chỉ chứa file
    này). ``mode=combined`` dùng toàn bộ ``documents`` của môn để tạo một kết quả
    tổng hợp duy nhất.
    """
    total = len(plan)
    independent = mode == "per_file" and focus is not None
    if independent:
        corpus = build_single_document_corpus(focus)
        context_label = "NỘI DUNG TÀI LIỆU DUY NHẤT"
        directive: str | None = SINGLE_DOCUMENT_DIRECTIVE
        cross_document = False
    else:
        corpus = build_subject_corpus(documents)
        context_label = "TOÀN BỘ NỘI DUNG TÀI LIỆU"
        directive = None
        cross_document = len(documents) > 1
    system_prompt = build_system_prompt(
        lang,
        plan=plan,
        cross_document=cross_document,
        single_document=independent,
    )
    user_content = build_document_user_prompt(
        corpus=corpus,
        plan=plan,
        subject=subject,
        documents=documents,
        focus_filename=focus.filename if focus else None,
        single_document=independent,
    )
    label = f"document:{focus.filename}" if focus else f"subject:{subject or 'all'}"
    filename = focus.filename if focus else (subject.strip() or "subject")

    logger.info(
        "Sinh nội dung cho %s (%s câu, %s tài liệu, ngôn ngữ %s)...",
        label,
        total,
        len(documents),
        lang,
    )

    try:
        payload, usages, model, raw_content = request_json_payload(
            system_prompt=system_prompt,
            user_content=user_content,
            label=label,
            # A curriculum must be one JSON object. json-repair runs inside
            # request_json_payload and broken answers are re-requested (2 retries).
            accept=(dict,),
        )
    except JsonPayloadError as exc:
        # Still useful for the client: hand back the metadata plus the raw text so
        # the frontend can show a friendly error and offer a retry.
        logger.error(
            "Không parse được JSON từ DeepSeek cho %s sau %s lần gọi.",
            label,
            exc.attempts,
        )
        return {
            "success": False,
            "filename": filename,
            "status_code": status.HTTP_502_BAD_GATEWAY,
            "model": exc.model,
            "error": "Model không trả về JSON hợp lệ.",
            "raw_response": exc.raw_content,
            "json_attempts": exc.attempts,
        }

    if not isinstance(payload, dict):
        logger.error("DeepSeek trả về JSON không phải object cho %s", label)
        return {
            "success": False,
            "filename": filename,
            "status_code": status.HTTP_502_BAD_GATEWAY,
            "model": model,
            "error": "Model không trả về đối tượng JSON hợp lệ.",
            "raw_response": raw_content[:2000],
        }

    questions, topup_usages, topup_rounds = complete_questions(
        questions=parse_questions(payload),
        plan=plan,
        system_prompt=system_prompt,
        corpus=corpus,
        label=label,
        context_label=context_label,
        subject=subject,
        focus_filename=focus.filename if focus else None,
        directive=directive,
    )
    usages.extend(topup_usages)

    summary = _as_text(payload.get("summary"))
    teaching = normalize_teaching(
        payload.get("teaching"), fallback_summary=summary, lang=lang
    )
    title = (
        _as_text(payload.get("title"))
        or subject.strip()
        or (Path(filename).stem if focus else "Tổng hợp môn học")
    )

    # Ở mode=combined số liệu là của cả môn; ở mode=per_file là của file tâm điểm.
    pages = sum(document.pages for document in documents) if focus is None else focus.pages
    characters = (
        sum(document.characters for document in documents)
        if focus is None
        else focus.characters
    )
    truncated = (
        any(document.truncated for document in documents)
        if focus is None
        else focus.truncated
    )
    formats = sorted({document.format for document in documents})
    document_format = (
        focus.format if focus else (formats[0] if len(formats) == 1 else "mixed")
    )

    return {
        "success": True,
        "filename": filename,
        "subject": subject.strip() or None,
        "mode": mode,
        "format": document_format,
        "model": model,
        "pages": pages,
        "characters_extracted": characters,
        "truncated": truncated,
        "documents_used": [document.filename for document in documents],
        "requested_count": total,
        "count": len(questions),
        "complete": len(questions) == total,
        "question_plan": question_plan_counts(plan),
        "topup_rounds": topup_rounds,
        "usage": merge_usage(usages),
        "data": {
            "title": title,
            "teaching": teaching,
            "questions": questions,
            "sources": document_sources(documents, focus.position if focus else None),
        },
    }


def parse_document_mode(mode: str) -> str:
    """Map the ``mode`` form field to ``per_file`` / ``combined``."""
    normalized = (mode or "").strip().lower().replace("-", "_")
    if normalized in {"", "per_file", "perfile", "each", "file"}:
        return "per_file"
    if normalized in {"combined", "single", "subject", "all"}:
        return "combined"
    raise HTTPException(
        status_code=status.HTTP_400_BAD_REQUEST,
        detail=(
            f"mode không hợp lệ: '{mode}'. "
            "Dùng 'per_file' (một kết quả mỗi file) hoặc 'combined' (một kết quả mỗi môn)."
        ),
    )


def clamp_document_count(count: int) -> int:
    """Clamp the requested question count to ``DOCUMENT_MIN/MAX_QUESTIONS``."""
    try:
        value = int(count)
    except (TypeError, ValueError):
        value = 0
    if value <= 0:
        value = DOCUMENT_DEFAULT_QUESTIONS
    return max(DOCUMENT_MIN_QUESTIONS, min(DOCUMENT_MAX_QUESTIONS, value))


# --------------------------------------------------------------------------------------
# Book splitting (POST /process-book)
# --------------------------------------------------------------------------------------
# Một quyển sách PDF được cắt thành từng chương theo mục lục (Inhaltsverzeichnis)
# rồi mỗi chương được xử lý bằng ĐÚNG pipeline của /process-document (teaching
# song ngữ Đức - Việt + bộ câu hỏi đủ số lượng). Thứ tự ưu tiên khi chia:
#   1. Mục lục: đọc các dòng "1 Grundlagen ... 5" trong trang Inhaltsverzeichnis.
#   2. Tiêu đề chương: dò "Kapitel 3 ..." / "3 Grundlagen" ở đầu các trang nội dung.
#   3. Chia đều theo trang: phương án cuối để endpoint luôn trả được kết quả.
BOOK_TOC_LABELS: tuple[str, ...] = (
    "inhaltsverzeichnis",
    "inhalt",
    "table of contents",
    "contents",
    "mục lục",
    "muc luc",
)
# Dòng mục lục: "1 Grundlagen ....... 12", "Kapitel 2: Recht ..... 30",
# "1. Einleitung .... 3". Chỉ nhận chương cấp 1 (1, 2, ...), bỏ mục con 2.1.
_TOC_ENTRY_RE = re.compile(
    r"^(?:(?:kapitel|chapter|teil|abschnitt|chương|chuong)\s+)?"
    r"(?P<number>\d{1,2})(?:\.(?!\d))?"
    r"[\s.:)–—-]+(?P<title>\S.*)$",
    re.IGNORECASE,
)
_TOC_TRAILING_PAGE_RE = re.compile(r"(?P<separator>[\s.·•…\-–—_]+)(?P<page>\d{1,4})\s*$")
# Tiêu đề chương ở đầu trang khi sách KHÔNG có mục lục (phương án 2).
_BOOK_HEADING_RE = re.compile(
    r"^(?:(?:kapitel|chapter|teil|abschnitt|chương|chuong)\s+)?"
    r"(?P<number>\d{1,2})(?:\.(?!\d))?"
    r"[\s.:)–—-]+(?P<title>[^\W\d_].{2,100})$",
    re.IGNORECASE | re.UNICODE,
)


@dataclass(slots=True)
class BookChapter:
    """Một chương của sách: mục lục (số/tên/trang in) + khoảng trang thật trong PDF."""

    index: int
    number: str
    title: str
    printed_page: int | None = None
    # Trang 1-based trong file PDF (KHÔNG phải số trang in trên mục lục).
    start_page: int = 0
    end_page: int = 0
    characters: int = 0

    @property
    def label(self) -> str:
        """Nhãn hiển thị, ví dụ ``"1 Grundlagen"``."""
        return " ".join(part for part in (self.number, self.title) if part).strip()


@dataclass(slots=True)
class BookSplit:
    """Kết quả chia sách: danh sách chương + cách tìm được chúng."""

    chapters: list[BookChapter]
    method: str  # "toc" | "headings" | "pages"
    toc_pages: tuple[int, int] | None = None  # 1-based, inclusive


def _normalize_toc_line(line: str) -> str:
    """Chuẩn hoá một dòng để so khớp tiêu đề mục lục (bỏ ký tự mềm, gộp khoảng trắng)."""
    return " ".join(line.replace("\u00ad", "").lower().split())


def _is_toc_heading(line: str) -> bool:
    """True nếu dòng là tiêu đề trang mục lục (Inhaltsverzeichnis / Table of contents)."""
    normalized = _normalize_toc_line(line)
    if not normalized:
        return False
    # "Inhaltsverzeichnis ..... 3" -> bỏ số trang và dấu chấm cuối dòng.
    cleaned = re.sub(r"[\s.·•…\-–—_0-9]+$", "", normalized).strip()
    for candidate in {normalized, cleaned}:
        for label in BOOK_TOC_LABELS:
            if candidate == label:
                return True
            if candidate.startswith(f"{label} ") and len(candidate) <= len(label) + 40:
                return True
    return False


def _split_toc_line(line: str, total_pages: int) -> tuple[str, str, int] | None:
    """Tách một dòng mục lục thành ``(số chương, tên chương, số trang in)``.

    Dòng không có số trang ở cuối, mục con (``2.1 ...``) hoặc số trang vô lý đều
    bị bỏ qua để không tạo ra chương ảo.
    """
    stripped = line.strip()
    if len(stripped) < 6:
        return None
    page_match = _TOC_TRAILING_PAGE_RE.search(stripped)
    if page_match is None:
        return None
    printed_page = int(page_match.group("page"))
    if printed_page <= 0 or printed_page > total_pages + BOOK_TOC_PAGE_TOLERANCE:
        return None
    head = stripped[: page_match.start()].strip().strip(".·•…-–—_ ")
    entry = _TOC_ENTRY_RE.match(head)
    if entry is None:
        return None
    title = " ".join(entry.group("title").split()).strip(".·•…-–—_ ")
    if len(title) < 3 or not re.search(r"[^\W\d_]", title, re.UNICODE):
        return None
    return entry.group("number"), title, printed_page


def _parse_toc_page(text: str, total_pages: int) -> list[BookChapter]:
    """Các dòng mục lục hợp lệ của MỘT trang (chưa biết trang bắt đầu nội dung)."""
    chapters: list[BookChapter] = []
    for line in text.splitlines():
        parsed = _split_toc_line(line, total_pages)
        if parsed is None:
            continue
        number, title, printed_page = parsed
        chapters.append(
            BookChapter(index=0, number=number, title=title, printed_page=printed_page)
        )
    return chapters


def _collect_toc(pages: Sequence[str], start_index: int) -> tuple[list[BookChapter], int]:
    """Đọc mục lục bắt đầu từ ``start_index``; trả về (chương, trang cuối mục lục)."""
    collected: list[BookChapter] = []
    seen: set[str] = set()
    end_index = start_index
    limit = min(len(pages), start_index + BOOK_TOC_MAX_PAGES)
    for page_index in range(start_index, limit):
        parsed = _parse_toc_page(pages[page_index], len(pages))
        if not parsed:
            if page_index == start_index:
                # Trang chỉ có tiêu đề "Inhaltsverzeichnis", mục lục nằm trang sau.
                continue
            break
        for chapter in parsed:
            if chapter.number in seen:
                continue
            seen.add(chapter.number)
            collected.append(chapter)
        end_index = page_index
    return collected, end_index


def _find_toc(pages: Sequence[str]) -> tuple[list[BookChapter], int | None, int | None]:
    """Tìm trang mục lục có nhiều dòng chương nhất trong phần đầu sách."""
    search_limit = min(len(pages), BOOK_TOC_SEARCH_PAGES)
    best_chapters: list[BookChapter] = []
    best_start: int | None = None
    best_end: int | None = None
    for index in range(search_limit):
        if not any(_is_toc_heading(line) for line in pages[index].splitlines()):
            continue
        chapters, end_index = _collect_toc(pages, index)
        if len(chapters) > len(best_chapters):
            best_chapters, best_start, best_end = chapters, index, end_index
    return best_chapters, best_start, best_end


def _heading_key(text: str) -> str:
    """Chuẩn hoá tiêu đề để so khớp dù PDF ngắt dòng / gạch nối / dấu câu."""
    lowered = text.replace("\u00ad", "").lower()
    lowered = re.sub(r"(?<=\w)[-\u2013\u2014]\s*\n?\s*(?=\w)", "", lowered)
    return re.sub(r"[\W_]+", "", lowered, flags=re.UNICODE)


def _number_pattern(number: str) -> re.Pattern[str]:
    """Regex nhận số chương đứng riêng (không khớp "1" trong "1.2" hay "10")."""
    return re.compile(
        rf"(?<![\d.,])(?:kapitel|chapter|teil|abschnitt|chương|chuong)?\s*"
        rf"{re.escape(number)}(?!\d)(?!\.\d)",
        re.IGNORECASE,
    )


def _page_starts_chapter(page_text: str, chapter: BookChapter) -> bool:
    """True nếu phần đầu trang là tiêu đề của ``chapter`` (trang bắt đầu chương)."""
    head = page_text[:1500]
    title_key = _heading_key(chapter.title)
    if not title_key:
        return False
    head_key = _heading_key(head)
    matched = title_key in head_key
    if not matched and len(title_key) >= 12:
        # Tiêu đề trong mục lục có thể dài hơn tiêu đề in trên trang.
        matched = title_key[:24] in head_key
    if not matched:
        return False
    if chapter.number and not _number_pattern(chapter.number).search(head):
        return False
    return True


def _find_chapter_start(
    pages: Sequence[str], chapter: BookChapter, start_index: int
) -> int | None:
    """Trang (0-based) đầu tiên từ ``start_index`` có tiêu đề của ``chapter``."""
    for index in range(max(0, start_index), len(pages)):
        if _page_starts_chapter(pages[index], chapter):
            return index
    return None


def _assign_chapter_pages(
    pages: Sequence[str], toc_chapters: Sequence[BookChapter], body_start: int
) -> list[BookChapter]:
    """Ghép mỗi mục lục với trang bắt đầu thật trong PDF rồi tính trang kết thúc."""
    if not pages:
        return []
    last_index = len(pages) - 1
    found: list[int | None] = []
    offsets: list[int] = []
    cursor = body_start
    for chapter in toc_chapters:
        start = _find_chapter_start(pages, chapter, cursor)
        found.append(start)
        if start is not None:
            cursor = start + 1
            if chapter.printed_page:
                offsets.append(start - (chapter.printed_page - 1))

    # Số trang in thường lệch số trang PDF (bìa, trang La Mã...): lấy offset phổ
    # biến nhất từ các chương đã dò được tiêu đề làm chuẩn cho chương còn lại.
    offset = Counter(offsets).most_common(1)[0][0] if offsets else 0

    assigned: list[BookChapter] = []
    previous = body_start - 1
    for chapter, start in zip(toc_chapters, found):
        page_index = start
        if page_index is None:
            page_index = (
                chapter.printed_page - 1 + offset
                if chapter.printed_page
                else previous + 1
            )
        page_index = max(body_start, min(last_index, page_index))
        if page_index <= previous:
            page_index = previous + 1
        if page_index > last_index:
            logger.warning(
                "Bỏ qua chương '%s': không còn trang nào sau chương trước.", chapter.label
            )
            continue
        assigned.append(
            BookChapter(
                index=len(assigned) + 1,
                number=chapter.number,
                title=chapter.title,
                printed_page=chapter.printed_page,
                start_page=page_index + 1,
            )
        )
        previous = page_index

    for position, chapter in enumerate(assigned):
        if position + 1 < len(assigned):
            end_index = assigned[position + 1].start_page - 2
        else:
            end_index = last_index
        chapter.end_page = max(chapter.start_page, end_index + 1)
    return assigned


def _detect_heading_pages(
    pages: Sequence[str], body_start: int
) -> list[tuple[int, str, str]]:
    """Dò tiêu đề chương (``Kapitel 3 ...`` / ``3 Grundlagen``) ở đầu mỗi trang."""
    detected: list[tuple[int, str, str]] = []
    seen_numbers: set[str] = set()
    for index in range(max(0, body_start), len(pages)):
        for line in pages[index].splitlines()[:5]:
            compact = " ".join(line.split())
            match = _BOOK_HEADING_RE.match(compact)
            if match is None:
                continue
            number = match.group("number")
            title = " ".join(match.group("title").split()).strip(".·•…-–—_ ")
            if number in seen_numbers or len(title) < 3:
                break
            seen_numbers.add(number)
            detected.append((index, number, title))
            break
    return detected


def _chapters_from_headings(
    pages: Sequence[str], headings: Sequence[tuple[int, str, str]]
) -> list[BookChapter]:
    """Dựng chương từ các trang tiêu đề dò được (khi sách không có mục lục)."""
    chapters: list[BookChapter] = []
    for position, (page_index, number, title) in enumerate(headings):
        start_page = page_index + 1
        end_page = (
            headings[position + 1][0] if position + 1 < len(headings) else len(pages)
        )
        chapters.append(
            BookChapter(
                index=len(chapters) + 1,
                number=number,
                title=title,
                start_page=start_page,
                end_page=max(start_page, end_page),
            )
        )
    return chapters


def _chunk_pages(pages: Sequence[str]) -> list[BookChapter]:
    """Phương án cuối: chia đều sách theo ``BOOK_FALLBACK_CHAPTER_PAGES`` trang."""
    size = BOOK_FALLBACK_CHAPTER_PAGES
    chapters: list[BookChapter] = []
    for start_index in range(0, len(pages), size):
        end_index = min(len(pages) - 1, start_index + size - 1)
        number = str(len(chapters) + 1)
        chapters.append(
            BookChapter(
                index=len(chapters) + 1,
                number=number,
                title=f"Phần {number} (trang {start_index + 1}-{end_index + 1})",
                start_page=start_index + 1,
                end_page=end_index + 1,
            )
        )
    return chapters


def split_book(pages: Sequence[str]) -> BookSplit:
    """Chia sách thành chương theo mục lục, tiêu đề chương hoặc chia đều theo trang.

    Đây là bước 2-3 của ``POST /process-book``: nội dung từng trang đã có sẵn (nhờ
    ``extract_pdf_pages``), hàm này chỉ quyết định ranh giới chương.
    """
    if not pages:
        return BookSplit(chapters=[], method="pages")

    toc_chapters, toc_start, toc_end = _find_toc(pages)
    if toc_chapters:
        body_start = (toc_end if toc_end is not None else toc_start or 0) + 1
        chapters = _assign_chapter_pages(pages, toc_chapters, body_start)
        if chapters:
            start = (toc_start or 0) + 1
            end = (toc_end or 0) + 1
            logger.info(
                "Mục lục trang %s-%s: %s chương (trang PDF bắt đầu: %s).",
                start,
                end,
                len(chapters),
                ", ".join(str(chapter.start_page) for chapter in chapters),
            )
            return BookSplit(chapters=chapters, method="toc", toc_pages=(start, end))
        logger.warning("Có mục lục nhưng không xác định được trang bắt đầu chương.")

    body_start = (toc_end + 1) if toc_end is not None else 0
    headings = _detect_heading_pages(pages, body_start)
    if len(headings) >= 2:
        chapters = _chapters_from_headings(pages, headings)
        logger.info("Không dùng được mục lục: dò được %s tiêu đề chương.", len(chapters))
        return BookSplit(chapters=chapters, method="headings")

    chapters = _chunk_pages(pages)
    logger.warning(
        "Không tìm thấy mục lục/tiêu đề chương: chia sách thành %s phần theo trang.",
        len(chapters),
    )
    return BookSplit(chapters=chapters, method="pages")


def clamp_book_chapter_limit(value: int) -> int:
    """Form ``max_chapters`` rỗng/0 -> ``MAX_BOOK_CHAPTERS``, kẹp trong 1..MAX."""
    try:
        requested = int(value)
    except (TypeError, ValueError):
        return MAX_BOOK_CHAPTERS
    if requested <= 0:
        return MAX_BOOK_CHAPTERS
    return max(1, min(MAX_BOOK_CHAPTERS, requested))


def resolve_book_background(value: str | None) -> bool:
    """Form ``background`` rỗng -> mặc định; ``false``/``0``/``no``/``off`` -> chạy đồng bộ."""
    if value is None or not value.strip():
        return BOOK_BACKGROUND_BY_DEFAULT
    return value.strip().lower() not in {"0", "false", "no", "off"}


def limit_book_chapters(
    chapters: Sequence[BookChapter], limit: int, total_pages: int
) -> tuple[list[BookChapter], bool]:
    """Giữ tối đa ``limit`` chương; chương cuối được nới tới hết sách để không mất nội dung.

    Mỗi chương tốn ít nhất một lần gọi DeepSeek, nên sách có quá nhiều chương phải
    gộp phần dư vào chương cuối thay vì âm thầm bỏ đi.
    """
    if limit <= 0 or len(chapters) <= limit:
        return list(chapters), False
    kept = [
        BookChapter(
            index=chapter.index,
            number=chapter.number,
            title=chapter.title,
            printed_page=chapter.printed_page,
            start_page=chapter.start_page,
            end_page=chapter.end_page,
        )
        for chapter in chapters[:limit]
    ]
    last = kept[-1]
    last.end_page = max(last.end_page, total_pages)
    last.title = f"{last.title} (và phần còn lại)".strip()
    return kept, True


def chapter_metadata(chapter: BookChapter) -> dict[str, Any]:
    """Metadata một chương trả về cho client (số trang in + trang PDF + số ký tự)."""
    return {
        "index": chapter.index,
        "number": chapter.number or None,
        "title": chapter.title,
        "label": chapter.label,
        "printed_page": chapter.printed_page,
        "start_page": chapter.start_page,
        "end_page": chapter.end_page,
        "pages": max(0, chapter.end_page - chapter.start_page + 1),
        "characters_extracted": chapter.characters,
    }


def build_chapter_document(
    chapter: BookChapter,
    pages: Sequence[str],
    *,
    book_stem: str,
    limit: int = MAX_TEXT_CHARS,
) -> ExtractedDocument:
    """Gom các trang của MỘT chương thành ``ExtractedDocument`` (giới hạn ``limit``)."""
    parts: list[str] = []
    for page_number in range(chapter.start_page, chapter.end_page + 1):
        page_text = pages[page_number - 1].strip()
        if page_text:
            parts.append(f"--- Trang {page_number} ---\n{page_text}")
    full_text = "\n\n".join(parts)
    chapter.characters = len(full_text)
    text = truncate_text(full_text, limit)
    if len(text) < len(full_text):
        logger.warning(
            "Chương '%s' có %s ký tự, cắt còn %s ký tự trước khi gửi DeepSeek.",
            chapter.label,
            len(full_text),
            len(text),
        )
    label = chapter.label or f"Chương {chapter.index}"
    return ExtractedDocument(
        position=chapter.index,
        filename=f"{book_stem} - {label}",
        text=text,
        pages=max(1, chapter.end_page - chapter.start_page + 1),
        characters=len(full_text),
        truncated=len(text) < len(full_text),
        format="pdf",
    )


def process_book_chapter(
    *,
    chapter: BookChapter,
    document: ExtractedDocument,
    subject: str,
    lang: str,
    plan: Sequence[str],
) -> dict[str, Any]:
    """Xử lý MỘT chương bằng đúng pipeline của ``/process-document`` (mode per_file).

    Mỗi chương chỉ gửi nội dung của chính nó cho DeepSeek, nên ``teaching`` và
    ``questions`` của phần tử kết quả không trộn kiến thức của chương khác.
    """
    logger.info(
        "Xử lý chương %s '%s': trang %s-%s, %s ký tự.",
        chapter.index,
        chapter.label,
        chapter.start_page,
        chapter.end_page,
        document.characters,
    )
    try:
        element = _generate_curriculum(
            documents=[document],
            focus=document,
            subject=subject,
            lang=lang,
            plan=plan,
            mode="per_file",
        )
    except QuestionShortfallError as exc:
        logger.warning("Chương %s: %s", chapter.label, exc)
        element = _error_element(
            document.filename,
            status.HTTP_502_BAD_GATEWAY,
            str(exc),
            requested_count=len(plan),
            count=exc.produced,
        )
    except HTTPException as exc:
        logger.warning("Chương %s lỗi %s: %s", chapter.label, exc.status_code, exc.detail)
        element = _error_element(document.filename, exc.status_code, str(exc.detail))
    except Exception as exc:  # noqa: BLE001 - một chương lỗi không làm hỏng cả sách
        logger.exception("Lỗi không mong đợi khi xử lý chương %s", chapter.label)
        element = _error_element(
            document.filename,
            status.HTTP_500_INTERNAL_SERVER_ERROR,
            f"Lỗi không mong đợi khi xử lý chương: {exc}",
        )
    element["chapter"] = chapter_metadata(chapter)
    return element


def process_book_chapters(
    chapters: Sequence[BookChapter],
    pages: Sequence[str],
    *,
    book_stem: str,
    subject: str,
    lang: str,
    plan: Sequence[str],
    on_progress: Callable[[list[dict[str, Any]]], None] | None = None,
) -> list[dict[str, Any]]:
    """Xử lý LẦN LƯỢT từng chương (tránh rate limit) và báo tiến độ sau mỗi chương."""
    results: list[dict[str, Any]] = []
    for chapter in chapters:
        document = build_chapter_document(chapter, pages, book_stem=book_stem)
        results.append(
            process_book_chapter(
                chapter=chapter,
                document=document,
                subject=subject,
                lang=lang,
                plan=plan,
            )
        )
        if on_progress is not None:
            on_progress(results)
    return results


# --- Job chạy ngầm (tránh timeout của Render với sách dày) ------------------------------
# Kết quả được giữ trong bộ nhớ process; Render chỉ chạy một instance nên đủ dùng,
# nhưng khởi động lại server sẽ xoá các job đang có (client gọi lại POST).
book_jobs: dict[str, dict[str, Any]] = {}
_book_jobs_lock = threading.Lock()


def _utc_now() -> str:
    """Thời điểm hiện tại dạng ISO-8601 (UTC)."""
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def _prune_book_jobs() -> None:
    """Xoá job quá TTL / quá số lượng để RAM không tăng mãi."""
    now = dt.datetime.now(dt.timezone.utc)
    with _book_jobs_lock:
        expired: list[str] = []
        for job_id, job in book_jobs.items():
            finished_at = job.get("finished_at")
            if not finished_at:
                continue
            try:
                finished = dt.datetime.fromisoformat(finished_at)
            except ValueError:
                expired.append(job_id)
                continue
            if (now - finished).total_seconds() >= BOOK_JOB_TTL_SECONDS:
                expired.append(job_id)
        for job_id in expired:
            book_jobs.pop(job_id, None)
        if len(book_jobs) > MAX_BOOK_JOBS:
            ordered = sorted(
                book_jobs.items(), key=lambda item: str(item[1].get("created_at", ""))
            )
            for job_id, _job in ordered[: len(book_jobs) - MAX_BOOK_JOBS]:
                book_jobs.pop(job_id, None)


def create_book_job(
    *,
    book: dict[str, Any],
    chapters: Sequence[BookChapter],
    request: dict[str, Any],
) -> dict[str, Any]:
    """Tạo job mới ở trạng thái ``queued`` và trả về bản sao JSON của nó."""
    _prune_book_jobs()
    job = {
        "job_id": uuid.uuid4().hex[:12],
        "status": "queued",
        "created_at": _utc_now(),
        "updated_at": _utc_now(),
        "finished_at": None,
        "book": book,
        "chapters": [chapter_metadata(chapter) for chapter in chapters],
        "progress": {"completed": 0, "total": len(chapters), "failed": 0},
        "results": [],
        "error": None,
        "request": request,
    }
    with _book_jobs_lock:
        book_jobs[job["job_id"]] = job
    return json.loads(json.dumps(job))


def update_book_job(job_id: str, **fields: Any) -> None:
    """Cập nhật một job (an toàn khi background thread và request cùng truy cập)."""
    with _book_jobs_lock:
        job = book_jobs.get(job_id)
        if job is None:
            return
        job.update(fields)
        job["updated_at"] = _utc_now()


def get_book_job(job_id: str) -> dict[str, Any] | None:
    """Bản sao JSON của job (``None`` nếu không tồn tại / đã hết TTL)."""
    _prune_book_jobs()
    with _book_jobs_lock:
        job = book_jobs.get(job_id)
        if job is None:
            return None
        return json.loads(json.dumps(job))


def run_book_job(
    job_id: str,
    chapters: Sequence[BookChapter],
    pages: Sequence[str],
    *,
    book_stem: str,
    subject: str,
    lang: str,
    plan: Sequence[str],
) -> None:
    """Hàm chạy ngầm: xử lý từng chương và cập nhật kết quả/tiến độ sau mỗi chương.

    ``BackgroundTasks`` gọi hàm này SAU khi response đã trả về, nên client không phải
    giữ kết nối trong lúc DeepSeek xử lý cả quyển sách (tránh timeout của Render).
    """
    update_book_job(job_id, status="running")
    try:

        def record(results: list[dict[str, Any]]) -> None:
            failed = sum(1 for element in results if not element.get("success"))
            update_book_job(
                job_id,
                results=list(results),
                progress={
                    "completed": len(results),
                    "total": len(chapters),
                    "failed": failed,
                },
            )

        results = process_book_chapters(
            chapters,
            pages,
            book_stem=book_stem,
            subject=subject,
            lang=lang,
            plan=plan,
            on_progress=record,
        )
    except Exception as exc:  # noqa: BLE001 - job phải dừng ở trạng thái failed
        logger.exception("Job %s thất bại", job_id)
        update_book_job(job_id, status="failed", error=str(exc), finished_at=_utc_now())
        return

    failed = sum(1 for element in results if not element.get("success"))
    update_book_job(
        job_id,
        status="completed",
        finished_at=_utc_now(),
        error=None,
        progress={"completed": len(results), "total": len(chapters), "failed": failed},
    )
    logger.info(
        "Job %s hoàn tất: %s/%s chương, %s lỗi.",
        job_id,
        len(results),
        len(chapters),
        failed,
    )


@app.post("/process-document", tags=["documents"])
async def process_document(
    files: Annotated[
        list[UploadFile],
        File(
            description=(
                "Danh sách tài liệu của cùng một môn — nhận .pdf, .docx, .xlsx, .txt, .csv "
                "(gửi lặp lại field 'files')"
            )
        ),
    ],
    lang: Annotated[
        str,
        Form(
            description=(
                "Ngôn ngữ hiển thị chính: 'de' hoặc 'vi' (mặc định: vi). "
                "teaching luôn có cả bản 'de' và 'vi' (summary_de/summary_vi, "
                "concepts_de/concepts_vi) bất kể giá trị này."
            )
        ),
    ] = DEFAULT_OUTPUT_LANGUAGE,
    mode: Annotated[
        str,
        Form(
            description=(
                "'per_file' (mặc định): mỗi file được trích xuất và phân tích ĐỘC LẬP "
                "(không gộp nội dung nhiều file); 'combined': một phần tử tổng hợp "
                "cho cả môn (tuỳ chọn, tương thích ngược)"
            )
        ),
    ] = "per_file",
    subject: Annotated[
        str,
        Form(description="Tên môn học, dùng cho tiêu đề và ngữ cảnh của kết quả"),
    ] = "",
    count: Annotated[
        int,
        Form(
            description=(
                f"Số câu hỏi mỗi kết quả ({DOCUMENT_MIN_QUESTIONS}-{DOCUMENT_MAX_QUESTIONS}, "
                "0 = mặc định)"
            )
        ),
    ] = 0,
    types: Annotated[
        str,
        Form(
            description=(
                "Các dạng câu cần tạo, phân tách bằng dấu phẩy; "
                "bỏ trống = đủ 4 dạng bắt buộc"
            )
        ),
    ] = "",
) -> list[dict[str, Any]]:
    """Nhận DANH SÁCH tài liệu của MỘT MÔN, xử lý TỪNG FILE ĐỘC LẬP, trả về MẢNG kết quả.

    - Field ``files`` gửi lặp lại cho mỗi file: ``files=a.pdf``, ``files=b.docx``, ...
      Định dạng được tự nhận diện theo đuôi file: ``.pdf`` (pypdf), ``.docx``
      (python-docx), ``.xlsx`` (openpyxl), ``.txt``, ``.csv`` (stdlib).
    - Server lặp qua từng file: trích xuất text của đúng file đó rồi gửi DUY NHẤT
      nội dung ấy cho DeepSeek. Prompt luôn nhấn mạnh: "Chỉ phân tích duy nhất tài
      liệu này. Không gộp nội dung với bất kỳ tài liệu nào khác. Tạo tóm tắt và
      khái niệm riêng cho chương này." Vì vậy mỗi phần tử của mảng kết quả chỉ
      phản ánh file tương ứng (``documents_used``/``data.sources`` chỉ có file đó).
    - ``lang``: ngôn ngữ hiển thị chính cho formulas/examples/questions; riêng
      ``teaching`` luôn song ngữ Đức - Việt: ``summary_de``/``summary_vi`` và
      ``concepts_de``/``concepts_vi``. Hai khoá cũ ``summary``/``key_concepts``
      vẫn được trả về như bản của ``lang`` cho client chưa kịp chuyển.
    - ``mode``: ``per_file`` (mặc định) trả về một phần tử cho mỗi file, mỗi phần
      tử là kết quả phân tích độc lập của file đó; ``combined`` (tuỳ chọn) trả về
      MỘT phần tử tổng hợp cho toàn môn.
    - ``subject``: tên môn học, dùng cho tiêu đề / ngữ cảnh.
    - ``count``: số câu hỏi mỗi kết quả, kẹp trong
      ``DOCUMENT_MIN_QUESTIONS``-``DOCUMENT_MAX_QUESTIONS``; ``0`` = mặc định.
    - ``types``: dạng câu cần tạo (``multiple_choice``, ``calculation``,
      ``fill_in_the_blank``, ``table``, ``multiple_response``, ``true_false``),
      phân tách bằng dấu phẩy. Bỏ trống = đủ 4 dạng bắt buộc, chia đều theo số câu.
    - Số câu được BẢO ĐẢM: nếu model trả thiếu, server gọi bổ sung (tối đa
      ``QUESTION_TOPUP_ATTEMPTS`` lần) và trả 502 nếu vẫn thiếu
      (``STRICT_QUESTION_COUNT``).
    - Một file lỗi (sai định dạng, quá lớn, file hỏng, DeepSeek lỗi...) KHÔNG làm
      hỏng cả request: phần tử tương ứng có ``success: false`` kèm
      ``status_code`` và ``error`` (415 định dạng không hỗ trợ, 400 file rỗng/sai
      định dạng thật, 413 quá lớn, 422 không đọc được nội dung, 503 server thiếu
      thư viện đọc Word/Excel).
    """
    if not files:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                "Vui lòng gửi ít nhất một tài liệu trong field 'files' "
                "(hỗ trợ .pdf, .docx, .xlsx, .txt, .csv)."
            ),
        )

    total = len(files)
    if total > MAX_BATCH_FILES:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=(
                f"Chỉ xử lý tối đa {MAX_BATCH_FILES} file mỗi request "
                f"(nhận được {total})."
            ),
        )

    document_mode = parse_document_mode(mode)
    requested_count = clamp_document_count(count)
    plan = plan_question_types(requested_count, types)
    resolved_lang = resolve_language(lang)
    subject_name = subject.strip()

    logger.info(
        "Bắt đầu xử lý %s tài liệu (mode=%s, môn=%s, %s câu, dạng=%s, ngôn ngữ: %s).",
        total,
        document_mode,
        subject_name or "-",
        requested_count,
        ",".join(question_plan_counts(plan)),
        resolved_lang,
    )

    # Thiếu API key là lỗi cấu hình của server: báo 500 ngay thay vì lặp qua
    # từng file rồi trả về hàng loạt phần tử lỗi khó hiểu cho client.
    get_deepseek_client()

    # Bước 1: trích xuất text của từng file. Mỗi file giữ nội dung riêng; ở bước 2
    # mỗi lần gọi DeepSeek chỉ nhận đúng nội dung của file đang xử lý.
    extracted: list[ExtractedDocument | None] = [None] * total
    errors: dict[int, dict[str, Any]] = {}
    for position, file in enumerate(files, start=1):
        filename = file.filename or f"document-{position}"
        logger.info("Trích xuất file %s/%s: %s", position, total, filename)
        try:
            extracted[position - 1] = await _extract_upload(file, position)
        except HTTPException as exc:
            logger.warning(
                "Bỏ qua file %s/%s (%s) do lỗi %s: %s",
                position,
                total,
                filename,
                exc.status_code,
                exc.detail,
            )
            errors[position] = _error_element(filename, exc.status_code, str(exc.detail))
        except Exception as exc:  # noqa: BLE001 - one bad file must not kill the batch
            logger.exception("Lỗi không mong đợi khi trích xuất %s", filename)
            errors[position] = _error_element(
                filename,
                status.HTTP_500_INTERNAL_SERVER_ERROR,
                f"Lỗi không mong đợi khi xử lý file: {exc}",
            )
        finally:
            # Đã đóng sau khi đọc, nhưng đóng lại vô hại và bao luôn các lỗi
            # xảy ra trước bước đọc (ví dụ sai content-type).
            await _close_upload(file)

    usable = [document for document in extracted if document is not None]
    if not usable:
        logger.info("Không file nào trích xuất được text, trả về toàn bộ phần tử lỗi.")
        return [
            errors.get(position)
            or _error_element(
                f"document-{position}",
                HTTP_422_UNPROCESSABLE_CONTENT,
                "Không đọc được file.",
            )
            for position in range(1, total + 1)
        ]

    # Bước 2: sinh nội dung. Chạy tuần tự để tránh vượt hạn mức (rate limit) và
    # giữ thứ tự kết quả ổn định.
    if document_mode == "combined":
        logger.info(
            "Sinh MỘT kết quả tổng hợp cho môn %s từ %s tài liệu.",
            subject_name or "-",
            len(usable),
        )
        combined_name = subject_name or "subject"
        try:
            element = _generate_curriculum(
                documents=usable,
                focus=None,
                subject=subject_name,
                lang=resolved_lang,
                plan=plan,
                mode=document_mode,
            )
        except QuestionShortfallError as exc:
            element = _error_element(
                combined_name,
                status.HTTP_502_BAD_GATEWAY,
                str(exc),
                requested_count=requested_count,
                count=exc.produced,
            )
        except HTTPException as exc:
            logger.warning("Tổng hợp môn %s lỗi %s: %s", combined_name, exc.status_code, exc.detail)
            element = _error_element(combined_name, exc.status_code, str(exc.detail))
        except Exception as exc:  # noqa: BLE001 - báo lỗi qua phần tử kết quả
            logger.exception("Lỗi không mong đợi khi tổng hợp môn %s", combined_name)
            element = _error_element(
                combined_name,
                status.HTTP_500_INTERNAL_SERVER_ERROR,
                f"Lỗi không mong đợi khi xử lý môn học: {exc}",
            )
        element["mode"] = document_mode
        element["failed_files"] = list(errors.values())
        logger.info(
            "Hoàn tất môn %s: %s câu hỏi, %s file lỗi.",
            combined_name,
            element.get("count", 0),
            len(errors),
        )
        return [element]

    results: list[dict[str, Any]] = []
    for position in range(1, total + 1):
        if position in errors:
            results.append(errors[position])
            continue
        document = extracted[position - 1]
        if document is None:  # phòng thủ: nhánh này không nên xảy ra
            results.append(
                _error_element(
                    f"document-{position}",
                    HTTP_422_UNPROCESSABLE_CONTENT,
                    "Không đọc được file.",
                )
            )
            continue
        logger.info("Sinh nội dung độc lập cho file %s/%s: %s", position, total, document.filename)
        try:
            results.append(
                _generate_curriculum(
                    # Chỉ file này được gửi cho DeepSeek: không gộp với file khác.
                    documents=[document],
                    focus=document,
                    subject=subject_name,
                    lang=resolved_lang,
                    plan=plan,
                    mode=document_mode,
                )
            )
        except QuestionShortfallError as exc:
            logger.warning("File %s: %s", document.filename, exc)
            results.append(
                _error_element(
                    document.filename,
                    status.HTTP_502_BAD_GATEWAY,
                    str(exc),
                    requested_count=requested_count,
                    count=exc.produced,
                )
            )
        except HTTPException as exc:
            logger.warning(
                "File %s/%s (%s) lỗi %s: %s",
                position,
                total,
                document.filename,
                exc.status_code,
                exc.detail,
            )
            results.append(_error_element(document.filename, exc.status_code, str(exc.detail)))
        except Exception as exc:  # noqa: BLE001 - one bad file must not kill the batch
            logger.exception("Lỗi không mong đợi khi xử lý %s", document.filename)
            results.append(
                _error_element(
                    document.filename,
                    status.HTTP_500_INTERNAL_SERVER_ERROR,
                    f"Lỗi không mong đợi khi xử lý file: {exc}",
                )
            )

    succeeded = sum(1 for result in results if result.get("success"))
    logger.info(
        "Hoàn tất %s file: %s thành công, %s lỗi.",
        total,
        succeeded,
        total - succeeded,
    )
    return results


@app.post("/process-book", tags=["books"])
async def process_book(
    background_tasks: BackgroundTasks,
    response: Response,
    file: Annotated[
        UploadFile,
        File(description="File PDF của CẢ quyển sách (mục lục nằm trong chính file này)"),
    ],
    lang: Annotated[
        str,
        Form(description="Ngôn ngữ hiển thị chính: 'de' hoặc 'vi' (mặc định: vi)"),
    ] = DEFAULT_OUTPUT_LANGUAGE,
    subject: Annotated[
        str,
        Form(description="Tên môn học, dùng cho tiêu đề và ngữ cảnh của từng chương"),
    ] = "",
    count: Annotated[
        int,
        Form(
            description=(
                f"Số câu hỏi MỖI chương ({DOCUMENT_MIN_QUESTIONS}-{DOCUMENT_MAX_QUESTIONS}, "
                "0 = mặc định)"
            )
        ),
    ] = 0,
    types: Annotated[
        str,
        Form(description="Dạng câu mỗi chương, phân tách bằng dấu phẩy; trống = 4 dạng bắt buộc"),
    ] = "",
    background: Annotated[
        str,
        Form(
            description=(
                "'true' (mặc định): trả 202 + job_id rồi xử lý ngầm (tránh timeout "
                "Render); 'false': trả thẳng MẢNG kết quả"
            )
        ),
    ] = "",
    max_chapters: Annotated[
        int,
        Form(
            description=(
                f"Số chương tối đa (1-{MAX_BOOK_CHAPTERS}, 0 = mặc định); "
                "phần dư gộp vào chương cuối"
            )
        ),
    ] = 0,
) -> Any:
    """Chia một quyển sách PDF thành từng chương rồi xử lý MỖI chương như /process-document.

    - ``file``: PDF của cả quyển sách. Sách được cắt theo mục lục
      (``Inhaltsverzeichnis`` / ``Table of contents``); nếu không tìm được mục lục
      thì dò tiêu đề chương (``Kapitel 3 ...``), cuối cùng là chia đều theo trang.
    - Mỗi chương được trích text riêng và gửi DUY NHẤT nội dung của chương đó cho
      DeepSeek qua đúng pipeline của ``/process-document`` (``mode=per_file``):
      ``teaching`` song ngữ Đức - Việt + đủ số câu hỏi đã yêu cầu.
    - Một chương lỗi KHÔNG làm hỏng cả sách: phần tử tương ứng có ``success: false``
      kèm ``status_code`` / ``error``, các chương còn lại vẫn được xử lý.
    - ``background`` mặc định ``true``: trả ngay ``202`` kèm ``job_id`` + danh sách
      chương rồi xử lý ngầm; client gọi ``GET /process-book/{job_id}`` để lấy tiến
      độ và kết quả nên không bị timeout khi sách dày. Gửi ``background=false`` để
      nhận thẳng mảng kết quả trong một response.
    - Mỗi phần tử kết quả có thêm ``chapter``: ``index``, ``number``, ``title``,
      ``label``, ``printed_page`` (số trang in trong mục lục), ``start_page`` /
      ``end_page`` (trang thật trong PDF), ``pages``, ``characters_extracted``.
    """
    filename = (file.filename or "book.pdf").strip() or "book.pdf"
    logger.info("Nhận sách %s để chia chương.", filename)

    try:
        data = await _read_upload(file, MAX_UPLOAD_BYTES)
    finally:
        await _close_upload(file)

    if not data:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"File {filename} rỗng, không có nội dung để xử lý.",
        )
    if not _looks_like_pdf(data):
        raise HTTPException(
            status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
            detail=(
                f"File {filename} không phải PDF hợp lệ. /process-book chỉ nhận file "
                ".pdf của cả quyển sách (mục lục nằm trong file)."
            ),
        )

    # Thiếu API key là lỗi cấu hình của server: báo ngay thay vì trả job lỗi.
    get_deepseek_client()

    requested_count = clamp_document_count(count)
    plan = plan_question_types(requested_count, types)
    resolved_lang = resolve_language(lang)
    subject_name = subject.strip()
    chapter_limit = clamp_book_chapter_limit(max_chapters)
    run_in_background = resolve_book_background(background)

    try:
        pages = extract_pdf_pages(data)
    except ValueError as exc:
        raise HTTPException(
            status_code=HTTP_422_UNPROCESSABLE_CONTENT, detail=str(exc)
        ) from exc

    split = split_book(pages)
    chapters, trimmed = limit_book_chapters(split.chapters, chapter_limit, len(pages))
    if not chapters:
        raise HTTPException(
            status_code=HTTP_422_UNPROCESSABLE_CONTENT,
            detail=(
                f"Không đọc được nội dung chương nào từ {filename} "
                "(PDF có thể là bản scan không có lớp text)."
            ),
        )

    book_stem = Path(filename).stem or "book"
    book_meta = {
        "filename": filename,
        "pages": len(pages),
        "characters_extracted": sum(len(page_text) for page_text in pages),
        "split_method": split.method,
        "toc_pages": list(split.toc_pages) if split.toc_pages else None,
        "chapters_detected": len(split.chapters),
        "chapters_processed": len(chapters),
        "chapters_trimmed": trimmed,
    }
    request_meta = {
        "subject": subject_name,
        "lang": resolved_lang,
        "count": requested_count,
        "types": question_plan_counts(plan),
        "background": run_in_background,
    }
    logger.info(
        "Sách %s: %s trang, chia bằng '%s' -> %s chương (xử lý %s, môn=%s, %s câu/chương, ngầm=%s).",
        filename,
        len(pages),
        split.method,
        len(split.chapters),
        len(chapters),
        subject_name or "-",
        requested_count,
        run_in_background,
    )

    if run_in_background:
        job = create_book_job(book=book_meta, chapters=chapters, request=request_meta)
        background_tasks.add_task(
            run_book_job,
            job["job_id"],
            chapters,
            pages,
            book_stem=book_stem,
            subject=subject_name,
            lang=resolved_lang,
            plan=plan,
        )
        response.status_code = status.HTTP_202_ACCEPTED
        return job

    results = process_book_chapters(
        chapters,
        pages,
        book_stem=book_stem,
        subject=subject_name,
        lang=resolved_lang,
        plan=plan,
    )
    succeeded = sum(1 for result in results if result.get("success"))
    logger.info(
        "Hoàn tất sách %s: %s/%s chương thành công.",
        filename,
        succeeded,
        len(results),
    )
    return results


@app.get("/process-book/{job_id}", tags=["books"])
async def get_process_book_job(job_id: str) -> dict[str, Any]:
    """Trạng thái + kết quả của job do ``POST /process-book`` tạo ra.

    - ``status``: ``queued`` -> ``running`` -> ``completed`` (hoặc ``failed``).
    - ``progress``: ``{completed, total, failed}`` — cập nhật sau MỖI chương, nên
      client poll vài giây một lần để hiển thị thanh tiến độ.
    - ``results``: mảng kết quả cùng cấu trúc với ``/process-document`` (kèm
      ``chapter``); phần tử được thêm dần trong lúc job chạy.
    - ``chapters``: danh sách chương đã chia (kèm số trang in / trang PDF).
    - Job nằm trong bộ nhớ process, bị xoá sau ``BOOK_JOB_TTL_SECONDS`` hoặc khi
      server khởi động lại (khi đó trả 404, client gọi lại ``POST /process-book``).
    """
    job = get_book_job(job_id)
    if job is None:
        raise HTTPException(
            status_code=status.HTTP_404_NOT_FOUND,
            detail=(
                f"Không tìm thấy job '{job_id}' (có thể đã hết hạn hoặc server vừa "
                "khởi động lại). Hãy gửi lại POST /process-book."
            ),
        )
    return job


@app.post("/generate-exam", tags=["exams"])
async def generate_exam(request: ExamRequest) -> dict[str, Any]:
    """Tạo đề luyện thi mới (câu hỏi trộn dạng) từ khuôn khổ kiến thức đã lưu.

    - ``context``: tóm tắt/khái niệm/công thức/ví dụ của một chương hoặc của mọi
      chương (client ghép sẵn). Nội dung được cắt theo ``MAX_EXAM_CONTEXT_CHARS``.
    - ``count``: số câu mong muốn, được kẹp trong ``EXAM_MIN/MAX_QUESTIONS``.
    - ``scope``: ``chapter`` (đề luyện tập) hoặc ``final`` (thi kết thúc học phần).
    - ``chapter_titles``: giúp model phân bổ câu hỏi đều cho mọi chương.
    - ``existing_questions``: các câu đã có, model phải tránh lặp lại.
    - ``seed``: mã biến thể do client sinh ra để mỗi lần bấm là một đề khác.

    Trả về ``{success, questions, count, model, usage, …}``. Danh sách câu hỏi
    giữ nguyên hợp đồng JSON của ``/process-document`` nên Next.js dùng chung
    ``parseCurriculum`` để chuẩn hoá.
    """
    context = request.context.strip()

    if not context:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Thiếu 'context': cần khuôn khổ kiến thức để tạo đề.",
        )

    count = clamp_exam_count(request.count)
    plan = plan_question_types(count, request.types)
    scope = request.scope if request.scope in EXAM_SCOPE_LABELS else "chapter"
    label = f"exam:{scope}"
    logger.info(
        "Tạo đề (%s): %s câu, dạng %s, %s chương, ngôn ngữ %s, seed %s.",
        scope,
        count,
        ",".join(question_plan_counts(plan)),
        len(request.chapter_titles),
        resolve_language(request.lang),
        request.seed,
    )

    system_prompt = build_exam_system_prompt(request.lang, plan)
    try:
        payload, usages, model, _raw_content = request_json_payload(
            system_prompt=system_prompt,
            user_content=build_exam_prompt(request, count, plan),
            label=label,
            temperature=EXAM_TEMPERATURE,
            max_tokens=EXAM_MAX_TOKENS,
            # The exam schema accepts the wrapper object or a bare question array.
            accept=(dict, list),
        )
    except JsonPayloadError as exc:
        logger.error(
            "Không parse được JSON đề thi sau %s lần gọi: %s",
            exc.attempts,
            exc.raw_content[:500],
        )
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Model không trả về JSON hợp lệ cho đề thi.",
        ) from exc

    # Accept both the documented wrapper and a bare array.
    first_questions = parse_questions(payload)
    if not first_questions:
        logger.warning("Đề %s: model không trả về câu hỏi hợp lệ, chuyển sang bổ sung.", scope)

    # Đề thi cũng phải đủ số câu và đúng phân bổ 4 dạng như /process-document.
    try:
        questions, topup_usages, topup_rounds = complete_questions(
            questions=first_questions,
            plan=plan,
            system_prompt=system_prompt,
            corpus=truncate_text(context, MAX_EXAM_CONTEXT_CHARS),
            label=label,
            context_label="KHUÔN KHỔ KIẾN THỨC",
            temperature=EXAM_TEMPERATURE,
            max_tokens=EXAM_MAX_TOKENS,
        )
    except QuestionShortfallError as exc:
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=str(exc),
        ) from exc

    usages.extend(topup_usages)
    logger.info(
        "Đề %s: %s/%s câu hỏi sau %s lần bổ sung.",
        scope,
        len(questions),
        count,
        topup_rounds,
    )

    return {
        "success": True,
        "scope": scope,
        "model": model,
        "requested_count": count,
        "count": len(questions),
        "complete": len(questions) == count,
        "question_plan": question_plan_counts(plan),
        "topup_rounds": topup_rounds,
        "seed": request.seed,
        "usage": merge_usage(usages),
        "questions": questions,
    }


@app.post("/explain-concept", tags=["concepts"])
async def explain_concept(request: ConceptRequest) -> dict[str, Any]:
    """Phân tích MỘT khái niệm theo đúng 4 phần: definition / erklaerung / zusammenhang / kritik.

    - ``concept``: tên khái niệm (bắt buộc); khoảng trắng thừa được gộp lại và
      cắt theo ``MAX_CONCEPT_NAME_CHARS``.
    - ``context``: nội dung tài liệu liên quan do client ghép sẵn, cắt theo
      ``MAX_CONCEPT_CONTEXT_CHARS``. Bỏ trống thì model dựa vào kiến thức chuẩn
      của môn học và ghi chú điều đó trong ``kritik``.
    - ``lang``: ``vi`` (mặc định) hoặc ``de`` — ngôn ngữ hiển thị chính; mỗi phần
      vẫn LUÔN có cả hai bản ``de`` và ``vi``.

    Trả về ``{success, concept, definition, erklaerung, zusammenhang, kritik,
    model, usage, …}``, trong đó 4 phần là object song ngữ
    ``{"de": "...", "vi": "..."}``. JSON hỏng được ``request_json_payload`` sửa
    bằng json-repair và gửi lại như các endpoint khác; vẫn hỏng thì báo 502.
    """
    # Gộp khoảng trắng để "  Khấu   hao " và "Khấu hao" cho cùng một kết quả/log.
    concept = " ".join(request.concept.split())
    if len(concept) > MAX_CONCEPT_NAME_CHARS:
        concept = concept[:MAX_CONCEPT_NAME_CHARS].rstrip()
    if not concept:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Thiếu 'concept': cần tên khái niệm cần phân tích.",
        )

    raw_context = request.context.strip()
    context = truncate_text(raw_context, MAX_CONCEPT_CONTEXT_CHARS)
    lang = resolve_language(request.lang)
    label = f"concept:{concept[:80]}"
    logger.info(
        "Giải thích khái niệm '%s' (%s/%s ký tự ngữ cảnh, ngôn ngữ %s).",
        concept,
        len(context),
        len(raw_context),
        lang,
    )

    try:
        payload, usages, model, _raw_content = request_json_payload(
            system_prompt=build_concept_system_prompt(lang),
            user_content=build_concept_user_prompt(concept=concept, context=context),
            label=label,
            temperature=EXPLAIN_TEMPERATURE,
            max_tokens=EXPLAIN_MAX_TOKENS,
        )
    except JsonPayloadError as exc:
        logger.error(
            "Không parse được JSON giải thích khái niệm sau %s lần gọi: %s",
            exc.attempts,
            exc.raw_content[:500],
        )
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail="Model không trả về JSON hợp lệ cho phần giải thích khái niệm.",
        ) from exc

    explanation = normalize_concept_explanation(payload, lang=lang)
    if not any(
        text
        for section in ("definition", "erklaerung")
        for text in explanation[section].values()
    ):
        received_keys = (
            ", ".join(sorted(str(key) for key in payload))
            if isinstance(payload, dict)
            else type(payload).__name__
        )
        logger.warning(
            "Khái niệm '%s': model không trả về definition/erklaerung (khoá nhận được: %s).",
            concept,
            received_keys,
        )
        raise HTTPException(
            status_code=status.HTTP_502_BAD_GATEWAY,
            detail=(
                "Model không trả về nội dung cho 'definition'/'erklaerung' "
                "(cả bản 'de' và 'vi')."
            ),
        )

    logger.info(
        "Đã giải thích khái niệm '%s' (%s ký tự definition tiếng Đức, %s ký tự tiếng Việt).",
        concept,
        len(explanation["definition"]["de"]),
        len(explanation["definition"]["vi"]),
    )

    return {
        "success": True,
        "concept": concept,
        "definition": explanation["definition"],
        "erklaerung": explanation["erklaerung"],
        "zusammenhang": explanation["zusammenhang"],
        "kritik": explanation["kritik"],
        "lang": lang,
        "model": model,
        "context_characters": len(context),
        "context_truncated": len(context) < len(raw_context),
        "usage": merge_usage(usages),
    }


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "main:app",
        host=os.getenv("HOST", "127.0.0.1"),
        port=int(os.getenv("PORT", "8000")),
        reload=True,
    )
