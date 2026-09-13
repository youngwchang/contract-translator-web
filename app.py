"""
contract-web / app.py  (v4)

v3 대비 변경
 - download_result 의 문법 오류 수정 (서버가 기동되지 않던 원인)
 - stdout 리디렉션을 스레드별로 분리 — 동시 작업 시 로그 섞임 방지
 - 백그라운드 태스크 참조 보관 — GC 로 인한 작업 유실 방지
 - 임시 파일을 성공 경로에서도 정리, mktemp 제거
 - SSE: 작업 소멸 감지, 클라이언트 끊김 감지, 프록시 유휴 종료 대비 keepalive
 - 완료 시 토큰 사용량 함께 전달
"""

import asyncio
import io
import json
import os
import secrets
import sys
import tempfile
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, StreamingResponse
from fastapi.staticfiles import StaticFiles

sys.path.insert(0, str(Path(__file__).parent))
import contract_translator as ct

MAX_UPLOAD_BYTES = 50 * 1024 * 1024
JOB_TTL_SEC      = 1800    # 완료된 작업 보관 시간
JOB_MAX_RUNTIME  = 3600    # 이 시간을 넘긴 running 작업은 실패 처리
SSE_PING_SEC     = 15

_executor = ThreadPoolExecutor(max_workers=4, thread_name_prefix="translate")
JOBS: dict[str, dict] = {}
_TASKS: set[asyncio.Task] = set()   # 태스크가 GC 되지 않도록 강참조 유지


# ── 스레드별 stdout 라우팅 ────────────────────────────────────────────────────
# sys.stdout 을 통째로 바꾸면 동시에 돌아가는 다른 작업의 print 까지 가로챈다.
# 라우터를 한 번만 설치하고, 스레드마다 어느 job 으로 보낼지 기록한다.
class _StdoutRouter(io.TextIOBase):
    def __init__(self, fallback):
        self._fallback = fallback
        self._local = threading.local()

    def bind(self, job_id: str):
        self._local.job_id = job_id
        self._local.buf = ""

    def unbind(self):
        self.flush()
        self._local.job_id = None

    def write(self, s: str) -> int:
        job_id = getattr(self._local, "job_id", None)
        if not job_id:
            return self._fallback.write(s)
        buf = getattr(self._local, "buf", "") + s
        while "\n" in buf:
            line, buf = buf.split("\n", 1)
            _push(job_id, line.strip())
        self._local.buf = buf
        return len(s)

    def flush(self):
        job_id = getattr(self._local, "job_id", None)
        buf = getattr(self._local, "buf", "")
        if job_id and buf.strip():
            _push(job_id, buf.strip())
        self._local.buf = ""
        try:
            self._fallback.flush()
        except Exception:
            pass


_router = _StdoutRouter(sys.__stdout__)
sys.stdout = _router


def _push(job_id: str, text: str, level: str = "info"):
    if not text:
        return
    job = JOBS.get(job_id)
    if job is None:          # 이미 정리된 작업 — 조용히 버린다
        return
    job["messages"].append({"type": "log", "level": level, "text": text})


def _run_in_thread(job_id: str, fn):
    _router.bind(job_id)
    try:
        return fn()
    finally:
        _router.unbind()


# ── 작업 정리 ─────────────────────────────────────────────────────────────────
def _cleanup_old_jobs():
    now = time.time()
    for jid, job in list(JOBS.items()):
        age = now - job["created_at"]
        if job["status"] in ("done", "error") and age > JOB_TTL_SEC:
            JOBS.pop(jid, None)
        elif job["status"] == "running" and age > JOB_MAX_RUNTIME:
            _push(jid, "❌ 오류: 처리 시간이 너무 길어 중단되었습니다.", "error")
            job["status"] = "error"


@asynccontextmanager
async def lifespan(_app: FastAPI):
    async def janitor():
        while True:
            await asyncio.sleep(300)
            try:
                _cleanup_old_jobs()
            except Exception:
                pass
    task = asyncio.create_task(janitor())
    try:
        yield
    finally:
        task.cancel()
        _executor.shutdown(wait=False, cancel_futures=True)


app = FastAPI(title="계약서 번역기", lifespan=lifespan)
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_methods=["*"], allow_headers=["*"])


# ── 헬스 체크 ──────────────────────────────────────────────────────────────────
@app.get("/health")
async def health():
    return {
        "status": "ok",
        "requiresCode": bool(os.environ.get("ACCESS_CODE")),
        "activeJobs": sum(1 for j in JOBS.values() if j["status"] == "running"),
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
    if required_code and not secrets.compare_digest(
            x_access_code.strip(), required_code.strip()):
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

    # 전량을 메모리에 올리기 전에 조각 단위로 읽으며 한도를 확인한다
    chunks, total = [], 0
    while True:
        chunk = await file.read(1024 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > MAX_UPLOAD_BYTES:
            raise HTTPException(413, "파일 크기는 50 MB 이하여야 합니다.")
        chunks.append(chunk)
    file_bytes = b"".join(chunks)
    if not file_bytes:
        raise HTTPException(400, "빈 파일입니다.")

    _cleanup_old_jobs()
    job_id = str(uuid.uuid4())
    JOBS[job_id] = {
        "status": "running", "messages": [],
        "file": None, "filename": "", "mime": "",
        "usage": None, "created_at": time.time(),
    }

    task = asyncio.create_task(
        run_translation(job_id, file_bytes, filename, ext, mode, fmt, api_key)
    )
    _TASKS.add(task)
    task.add_done_callback(_TASKS.discard)
    return {"job_id": job_id}


def _sse(payload: dict) -> str:
    return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"


# ── SSE 스트리밍 ───────────────────────────────────────────────────────────────
@app.get("/api/stream/{job_id}")
async def stream_progress(job_id: str, request: Request):
    if job_id not in JOBS:
        raise HTTPException(404, "작업을 찾을 수 없습니다.")

    async def generator():
        sent = 0
        last_out = time.monotonic()
        while True:
            if await request.is_disconnected():
                return

            job = JOBS.get(job_id)
            if job is None:
                yield _sse({"type": "error", "text": "작업 정보가 만료되었습니다."})
                return

            msgs = job["messages"]
            while sent < len(msgs):
                yield _sse(msgs[sent])
                sent += 1
                last_out = time.monotonic()

            status = job["status"]
            if status in ("done", "error"):
                payload: dict = {"type": status}
                if status == "done":
                    payload["filename"] = job.get("filename", "result.xlsx")
                    payload["usage"] = job.get("usage")
                yield _sse(payload)
                return

            # 프록시가 유휴 커넥션을 끊지 않도록 주석 프레임을 보낸다
            if time.monotonic() - last_out > SSE_PING_SEC:
                yield ": ping\n\n"
                last_out = time.monotonic()

            await asyncio.sleep(0.4)

    return StreamingResponse(
        generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no",
                 "Connection": "keep-alive"},
    )


# ── 파일 다운로드 ──────────────────────────────────────────────────────────────
@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(404, f"job_id '{job_id}' 없음. 서버가 재시작되었을 수 있습니다.")
    if job["status"] != "done":
        raise HTTPException(400, f"작업 상태: {job['status']}")
    if not job.get("file"):
        raise HTTPException(404, "파일 데이터가 없습니다.")

    mime = job.get("mime",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    name = job["filename"]
    safe = quote(name, safe="")
    ascii_fallback = f"contract_result{Path(name).suffix or '.xlsx'}"

    return Response(
        content=job["file"],
        media_type=mime,
        headers={
            "Content-Disposition": (
                f'attachment; filename="{ascii_fallback}"; '
                f"filename*=UTF-8''{safe}"
            ),
        },
    )


# ── 번역 실행 (백그라운드 태스크) ─────────────────────────────────────────────
async def run_translation(
    job_id: str, file_bytes: bytes, filename: str,
    ext: str, mode: str, fmt: str, api_key: str,
):
    loop = asyncio.get_running_loop()

    def push(msg: str, level: str = "info"):
        _push(job_id, msg, level)

    async def off_thread(fn):
        return await loop.run_in_executor(
            _executor, lambda: _run_in_thread(job_id, fn)
        )

    tmp_path = None
    out_path = None
    try:
        push("📄 파일 수신 완료 — 텍스트 추출 중...")
        ct.reset_usage()

        with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = Path(tmp.name)

        # 1. 텍스트 추출
        text: str = await off_thread(lambda: ct.extract_text(tmp_path))

        # 2. 이미지 PDF 감지 → OCR
        if ext == ".pdf" and len(text.strip()) < 300:
            push("🔍 이미지(스캔) PDF 감지 — Claude Vision OCR 시작", "warn")
            push("⏳ 페이지 수에 따라 수분 소요될 수 있습니다...", "warn")
            import anthropic as _ant
            client_ocr = _ant.Anthropic(api_key=api_key)
            text = await off_thread(lambda: ct.extract_pdf_ocr(tmp_path, client_ocr))

        if len(text.strip()) < 100:
            raise ValueError("텍스트를 추출할 수 없습니다. 파일을 확인해주세요.")

        char_count = len(text)
        est_pages  = max(1, char_count // 600)
        push(f"✅ 텍스트 추출 완료 — {char_count:,}자 (약 {est_pages}페이지 추정)")

        # 3. Claude API 번역
        import anthropic as _ant
        client     = _ant.Anthropic(api_key=api_key)
        mode_label = "핵심 조항 분석" if mode == "check" else "전체 번역"
        push(f"🤖 Claude {mode_label} 시작...")
        if mode == "full" and est_pages > 30:
            push(f"   ※ {est_pages}페이지 분량 — 여러 번에 나눠 처리합니다", "warn")

        if mode == "check":
            results = await off_thread(lambda: ct.analyze_key_terms(client, text))
        else:
            results = await off_thread(lambda: ct.translate_contract(client, text))

        if not results:
            raise ValueError("조항을 추출하지 못했습니다. 파일 내용을 확인해주세요.")

        fmt_label = "Word" if fmt == "word" else "Excel"
        push(f"📊 {len(results)}개 조항 처리 완료 — {fmt_label} 파일 생성 중...")

        # 4. 파일 저장
        date_str   = datetime.now().strftime("%Y%m%d_%H%M")
        suffix_str = "_핵심조항" if mode == "check" else "_전체번역"

        if fmt == "word":
            out_suffix = ".docx"
            mime_type  = ("application/vnd.openxmlformats-officedocument"
                          ".wordprocessingml.document")
            save_fn    = ct.save_key_terms_word if mode == "check" else ct.save_word
        else:
            out_suffix = ".xlsx"
            mime_type  = ("application/vnd.openxmlformats-officedocument"
                          ".spreadsheetml.sheet")
            save_fn    = ct.save_key_terms_excel if mode == "check" else ct.save_excel

        out_name = f"{Path(filename).stem}{suffix_str}_{date_str}{out_suffix}"
        with tempfile.NamedTemporaryFile(suffix=out_suffix, delete=False) as out_tmp:
            out_path = Path(out_tmp.name)

        await off_thread(lambda: save_fn(results, out_path, filename))

        out_bytes = out_path.read_bytes()
        usage = ct.get_usage()

        JOBS[job_id].update({
            "file": out_bytes, "filename": out_name,
            "mime": mime_type, "usage": usage, "status": "done",
        })
        push(f"✅ 완료! {out_name}  ({len(out_bytes) / 1024:.0f} KB)")
        if usage.get("calls"):
            push(f"   토큰 — 입력 {usage['input']:,} / 출력 {usage['output']:,}"
                 f" / 캐시 적중 {usage['cache_read']:,} (호출 {usage['calls']}회)")

    except Exception as exc:
        _push(job_id, f"❌ 오류: {type(exc).__name__}: {exc}", "error")
        job = JOBS.get(job_id)
        if job:
            job["status"] = "error"
    finally:
        for p in (tmp_path, out_path):
            if p:
                try:
                    p.unlink(missing_ok=True)
                except Exception:
                    pass


# ── 정적 파일 서빙 ─────────────────────────────────────────────────────────────
public_dir = Path(__file__).parent / "public"
public_dir.mkdir(exist_ok=True)
app.mount("/", StaticFiles(directory=str(public_dir), html=True), name="static")
