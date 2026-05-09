"""
contract-web / app.py  (v3 — 한글 파일명 다운로드 수정)
"""

import asyncio
import io
import json
import os
import sys
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from fastapi.staticfiles import StaticFiles

sys.path.insert(0, str(Path(__file__).parent))
import contract_translator as ct

app = FastAPI(title="계약서 번역기")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_methods=["*"], allow_headers=["*"])

_executor = ThreadPoolExecutor(max_workers=4)
JOBS: dict[str, dict] = {}


def _cleanup_old_jobs():
    now = time.time()
    stale = [jid for jid, j in list(JOBS.items())
             if j["status"] in ("done", "error") and now - j["created_at"] > 1800]
    for jid in stale:
        del JOBS[jid]


# ── 스레드별 stdout 리디렉션 ───────────────────────────────────────────────────
class _JobStream(io.TextIOBase):
    def __init__(self, job_id: str):
        self.job_id = job_id
        self._buf   = ""

    def write(self, s: str) -> int:
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.strip()
            if line:
                JOBS[self.job_id]["messages"].append(
                    {"type": "log", "level": "info", "text": line}
                )
        return len(s)

    def flush(self):
        if self._buf.strip():
            JOBS[self.job_id]["messages"].append(
                {"type": "log", "level": "info", "text": self._buf.strip()}
            )
            self._buf = ""


def _run_in_thread(job_id: str, fn):
    stream     = _JobStream(job_id)
    old_stdout = sys.stdout
    sys.stdout = stream
    try:
        return fn()
    finally:
        stream.flush()
        sys.stdout = old_stdout


# ── 헬스 체크 ──────────────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {
        "status": "ok",
        "requiresCode": bool(os.environ.get("ACCESS_CODE")),
    }


# ── 번역 작업 시작 ─────────────────────────────────────────────────────────────
@app.post("/api/translate")
async def start_translate(
    file: UploadFile = File(...),
    mode: str = Form("full"),
    fmt:  str = Form("excel"),
    x_access_code: str = Header(default=""),
):
    required_code = os.environ.get("ACCESS_CODE", "")
    if required_code and x_access_code.strip() != required_code.strip():
        raise HTTPException(403, "접근 코드가 올바르지 않습니다.")

    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        raise HTTPException(500, "서버에 ANTHROPIC_API_KEY가 설정되지 않았습니다.")

    filename = file.filename or "contract"
    ext = Path(filename).suffix.lower()
    if ext not in (".pdf", ".docx"):
        raise HTTPException(400, "PDF 또는 DOCX 파일만 지원합니다.")
    if mode not in ("full", "check"):
        raise HTTPException(400, "mode 는 'full' 또는 'check' 여야 합니다.")
    if fmt not in ("excel", "word"):
        raise HTTPException(400, "fmt 는 'excel' 또는 'word' 여야 합니다.")

    file_bytes = await file.read()
    if len(file_bytes) > 50 * 1024 * 1024:
        raise HTTPException(413, "파일 크기는 50 MB 이하여야 합니다.")

    _cleanup_old_jobs()
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status": "running", "messages": [],
        "file": None, "filename": "", "mime": "",
        "created_at": time.time(),
    }

    asyncio.create_task(
        run_translation(job_id, file_bytes, filename, ext, mode, fmt, api_key)
    )
    return {"job_id": job_id}


# ── SSE 스트리밍 ───────────────────────────────────────────────────────────────
@app.get("/api/stream/{job_id}")
async def stream_progress(job_id: str):
    if job_id not in JOBS:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")

    async def generator():
        sent = 0
        while True:
            job  = JOBS.get(job_id, {})
            msgs = job.get("messages", [])
            while sent < len(msgs):
                yield f"data: {json.dumps(msgs[sent], ensure_ascii=False)}\n\n"
                sent += 1
            status = job.get("status")
            if status in ("done", "error"):
                payload: dict = {"type": status}
                if status == "done":
                    payload["filename"] = job.get("filename", "result.xlsx")
                yield f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"
                break
            await asyncio.sleep(0.4)

    return StreamingResponse(
        generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


# ── 파일 다운로드 ──────────────────────────────────────────────────────────────
@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    job = JOBS.get(job_id)
    if not job or job["status"] != "done" or not job.get("file"):
        raise HTTPException(404, "다운로드할 파일이 없습니다.")

    mime      = job.get("mime", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    safe_name = quote(job["filename"], safe="")   # 한글 파일명 RFC 5987 인코딩

    return StreamingResponse(
        io.BytesIO(job["file"]),
        media_type=mime,
        headers={
            "Content-Disposition": f"attachment; filename*=UTF-8''{safe_name}"
        },
    )


# ── 번역 실행 (백그라운드 태스크) ─────────────────────────────────────────────
async def run_translation(
    job_id: str, file_bytes: bytes, filename: str,
    ext: str, mode: str, fmt: str, api_key: str,
):
    loop = asyncio.get_event_loop()

    def push(msg: str, level: str = "info"):
        JOBS[job_id]["messages"].append({"type": "log", "level": level, "text": msg})

    tmp_path = None
    try:
        push("📄 파일 수신 완료 — 텍스트 추출 중...")

        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = Path(tmp.name)

        # 1. 텍스트 추출
        text: str = await loop.run_in_executor(
            _executor,
            lambda: _run_in_thread(job_id, lambda: ct.extract_text(tmp_path))
        )

        # 2. 이미지 PDF 감지 → OCR
        is_image_pdf = (ext == ".pdf" and len(text.strip()) < 300)
        if is_image_pdf:
            push("🔍 이미지(스캔) PDF 감지 — Claude Vision OCR 시작", "warn")
            push("⏳ 페이지 수에 따라 수분 소요될 수 있습니다...", "warn")
            import anthropic as _ant
            client_ocr = _ant.Anthropic(api_key=api_key)
            text = await loop.run_in_executor(
                _executor,
                lambda: _run_in_thread(job_id, lambda: ct.extract_pdf_ocr(tmp_path, client_ocr))
            )

        if len(text.strip()) < 100:
            raise ValueError("텍스트를 추출할 수 없습니다. 파일을 확인해주세요.")

        char_count = len(text)
        est_pages  = max(1, char_count // 600)
        push(f"✅ 텍스트 추출 완료 — {char_count:,}자 (약 {est_pages}페이지 추정)")

        # 3. Claude API 번역
        import anthropic as _ant
        client     = _ant.Anthropic(api_key=api_key)
        mode_label = "핵심 조항 분석" if mode == "check" else "전체 번역"
        push(f"🤖 Claude AI {mode_label} 시작...")
        if est_pages > 30:
            push(f"   ※ {est_pages}페이지 분량 — 청크 분할 처리로 시간이 걸립니다", "warn")

        if mode == "check":
            results = await loop.run_in_executor(
                _executor,
                lambda: _run_in_thread(job_id, lambda: ct.analyze_key_terms(client, text))
            )
        else:
            results = await loop.run_in_executor(
                _executor,
                lambda: _run_in_thread(job_id, lambda: ct.translate_contract(client, text))
            )

        if not results:
            raise ValueError("조항을 추출하지 못했습니다. 파일 내용을 확인해주세요.")

        fmt_label = "Word" if fmt == "word" else "Excel"
        push(f"📊 번역 완료 — {len(results)}개 조항. {fmt_label} 파일 생성 중...")

        # 4. 파일 저장
        date_str   = datetime.now().strftime("%Y%m%d_%H%M")
        suffix_str = "_핵심조항" if mode == "check" else "_전체번역"

        if fmt == "word":
            out_suffix = ".docx"
            mime_type  = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        else:
            out_suffix = ".xlsx"
            mime_type  = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        out_name = f"{Path(filename).stem}{suffix_str}_{date_str}{out_suffix}"
        out_path = Path(tempfile.mktemp(suffix=out_suffix))

        if fmt == "word":
            save_fn = ct.save_key_terms_word if mode == "check" else ct.save_word
        else:
            save_fn = ct.save_key_terms_excel if mode == "check" else ct.save_excel

        await loop.run_in_executor(
            _executor,
            lambda: _run_in_thread(job_id, lambda: save_fn(results, out_path, filename))
        )

        out_bytes = out_path.read_bytes()
        out_path.unlink(missing_ok=True)

        JOBS[job_id].update({
            "file": out_bytes, "filename": out_name,
            "mime": mime_type, "status": "done",
        })
        push(f"✅ 완료! {out_name}  ({len(out_bytes) // 1024:.0f} KB)")

    except Exception as exc:
        if tmp_path:
            tmp_path.unlink(missing_ok=True)
        JOBS[job_id]["messages"].append(
            {"type": "log", "level": "error", "text": f"❌ 오류: {exc}"}
        )
        JOBS[job_id]["status"] = "error"


# ── 정적 파일 서빙 ─────────────────────────────────────────────────────────────
public_dir = Path(__file__).parent / "public"
public_dir.mkdir(exist_ok=True)
app.mount("/", StaticFiles(directory=str(public_dir), html=True), name="static")
