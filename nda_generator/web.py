"""Application web FastAPI : upload NDA + playbook, suivi par issue (SSE), téléchargement du DOCX."""

from __future__ import annotations

import io
import json
import logging
import queue
import shutil
import tempfile
import threading
import uuid
from contextlib import asynccontextmanager
from pathlib import Path

from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, PlainTextResponse, Response, StreamingResponse
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

import mammoth

from nda_generator.docx_preview_html import docx_revision_html_fragment
from nda_generator.pipeline import run_review
from nda_generator.playbook import load_playbook

_JOBS_LOCK = threading.Lock()
_JOBS: dict[str, dict] = {}

STATIC_DIR = Path(__file__).resolve().parent / "static"


@asynccontextmanager
async def lifespan(_app: FastAPI):
    load_dotenv()
    yield


app = FastAPI(title="NDA Generator", lifespan=lifespan)
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


@app.get("/")
async def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


def _setup_job_logger(job: dict) -> logging.Logger:
    log_buffer: io.StringIO = job["log_buffer"]
    log = logging.getLogger(f"nda_job.{job['job_id']}")
    log.handlers.clear()
    log.setLevel(logging.DEBUG if job["verbose"] else logging.INFO)
    log.propagate = False
    h = logging.StreamHandler(log_buffer)
    h.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    log.addHandler(h)
    return log


def _run_job_worker(job: dict) -> None:
    log = _setup_job_logger(job)
    last_pct = 0

    def on_progress(ev: dict) -> None:
        nonlocal last_pct
        if isinstance(ev.get("percent"), (int, float)):
            last_pct = int(ev["percent"])
        job["event_q"].put(ev)

    try:
        ok = run_review(
            nda_path=job["nda_path"],
            playbook_path=job["pb_path"],
            out_path=job["out_path"],
            author=job["author"],
            strict_ops=job["strict_ops"],
            log=log,
            on_progress=on_progress,
        )
        job["success"] = ok
    except Exception:
        logging.getLogger("nda_job").exception("Job %s", job["job_id"])
        job["success"] = False
    finally:
        job["event_q"].put(
            {
                "kind": "complete",
                "success": bool(job.get("success")),
                "percent": 100 if job.get("success") else last_pct,
            }
        )


@app.post("/api/jobs")
async def create_job(
    nda: UploadFile = File(...),
    playbook: UploadFile = File(...),
    author: str = Form(""),
    strict_ops: str = Form(""),
    verbose: str = Form(""),
) -> JSONResponse:
    if not nda.filename or not nda.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Le NDA doit être un fichier .docx")
    if not playbook.filename or not playbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Le playbook doit être un fichier .xlsx")

    content = await nda.read()
    if not content:
        raise HTTPException(status_code=400, detail="Fichier NDA vide")
    pb_bytes = await playbook.read()
    if not pb_bytes:
        raise HTTPException(status_code=400, detail="Fichier playbook vide")

    job_id = uuid.uuid4().hex
    work_dir = Path(tempfile.mkdtemp(prefix=f"nda_job_{job_id}_"))
    nda_path = work_dir / "nda.docx"
    pb_path = work_dir / "playbook.xlsx"
    out_path = work_dir / "out.docx"
    nda_path.write_bytes(content)
    pb_path.write_bytes(pb_bytes)

    try:
        issues_objs = load_playbook(pb_path)
    except Exception as e:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail=f"Playbook invalide : {e}") from e

    if not issues_objs:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=400, detail="Aucune issue dans le playbook")

    job = {
        "job_id": job_id,
        "work_dir": work_dir,
        "nda_path": nda_path,
        "pb_path": pb_path,
        "out_path": out_path,
        "author": author,
        "strict_ops": strict_ops == "1",
        "verbose": verbose == "1",
        "log_buffer": io.StringIO(),
        "event_q": queue.Queue(),
        "started": False,
        "start_lock": threading.Lock(),
        "success": None,
    }
    with _JOBS_LOCK:
        _JOBS[job_id] = job

    issues_payload = [{"index": i + 1, "title": x.nom} for i, x in enumerate(issues_objs)]
    return JSONResponse({"job_id": job_id, "issues": [x["title"] for x in issues_payload]})


@app.get("/api/jobs/{job_id}/events")
async def job_events(job_id: str) -> StreamingResponse:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job introuvable")

    def start_worker_once() -> None:
        with job["start_lock"]:
            if job["started"]:
                return
            job["started"] = True
            threading.Thread(target=_run_job_worker, args=(job,), daemon=True).start()

    async def event_gen():
        start_worker_once()
        while True:
            item: dict = await run_in_threadpool(job["event_q"].get)
            yield f"data: {json.dumps(item, ensure_ascii=False)}\n\n"
            if item.get("kind") == "complete":
                break

    return StreamingResponse(
        event_gen(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


def _docx_to_preview_html(path: Path) -> str:
    data = path.read_bytes()
    try:
        inner = docx_revision_html_fragment(data)
        if not inner or not inner.strip():
            raise ValueError("aperçu vide")
    except Exception:
        result = mammoth.convert_to_html(io.BytesIO(data))
        inner = result.value or "<p>(Document vide.)</p>"
    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body {{
      font-family: "Georgia", "Times New Roman", serif;
      font-size: 15px;
      line-height: 1.55;
      color: #1a1a1a;
      background: #faf9f7;
      margin: 0;
      padding: 1.25rem 1.5rem 2rem;
      max-width: 42rem;
    }}
    p {{ margin: 0.65em 0; }}
    h1, h2, h3, h4, h5, h6 {{
      font-family: system-ui, "Segoe UI", "Helvetica Neue", Arial, sans-serif;
      font-weight: 700;
      line-height: 1.25;
      color: #111;
      letter-spacing: -0.02em;
    }}
    h1 {{
      font-size: 1.65rem;
      margin: 1.35em 0 0.45em;
      font-weight: 700;
      border-bottom: 1px solid #ddd;
      padding-bottom: 0.25em;
    }}
    h2 {{ font-size: 1.35rem; margin: 1.15em 0 0.4em; font-weight: 650; }}
    h3 {{ font-size: 1.15rem; margin: 1em 0 0.35em; font-weight: 650; }}
    h4 {{ font-size: 1.05rem; margin: 0.9em 0 0.3em; font-weight: 600; }}
    h5 {{ font-size: 0.98rem; margin: 0.85em 0 0.28em; font-weight: 600; text-transform: uppercase; letter-spacing: 0.04em; }}
    h6 {{ font-size: 0.95rem; margin: 0.8em 0 0.25em; font-weight: 600; color: #333; }}
    table {{ border-collapse: collapse; width: 100%; margin: 1em 0; }}
    td, th {{ border: 1px solid #ccc; padding: 0.35rem 0.5rem; vertical-align: top; }}
    .rev-del {{
      background-color: rgba(220, 38, 38, 0.32);
      color: #7f1d1d;
      text-decoration: line-through;
      text-decoration-color: #b91c1c;
      text-decoration-thickness: 0.07em;
      border-radius: 2px;
      padding: 0 0.12em;
    }}
    .rev-ins {{
      background-color: rgba(22, 163, 74, 0.35);
      color: #14532d;
      border-radius: 2px;
      padding: 0 0.12em;
    }}
  </style>
</head>
<body>
{inner}
</body>
</html>"""


@app.get("/api/jobs/{job_id}/preview")
async def job_preview(job_id: str) -> HTMLResponse:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job introuvable")
    if not job.get("success"):
        raise HTTPException(status_code=400, detail="La revue n’a pas réussi ou n’est pas terminée")
    out_path: Path = job["out_path"]
    if not out_path.is_file():
        raise HTTPException(status_code=404, detail="Fichier de sortie absent")
    html = await run_in_threadpool(_docx_to_preview_html, out_path)
    return HTMLResponse(content=html)


@app.get("/api/jobs/{job_id}/log")
async def job_log(job_id: str) -> PlainTextResponse:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job introuvable")
    return PlainTextResponse(job["log_buffer"].getvalue())


@app.get("/api/jobs/{job_id}/download")
async def job_download(job_id: str) -> Response:
    with _JOBS_LOCK:
        job = _JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job introuvable")
    if not job.get("success"):
        raise HTTPException(status_code=400, detail="La revue n’a pas réussi ou n’est pas terminée")
    out_path: Path = job["out_path"]
    if not out_path.is_file():
        raise HTTPException(status_code=404, detail="Fichier de sortie absent")

    data = out_path.read_bytes()
    work_dir: Path = job["work_dir"]
    with _JOBS_LOCK:
        _JOBS.pop(job_id, None)
    shutil.rmtree(work_dir, ignore_errors=True)

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="NDA_revu.docx"'},
    )


def _run_review_sync(
    nda_path: Path,
    playbook_path: Path,
    out_path: Path,
    author: str,
    strict_ops: bool,
    log: logging.Logger,
) -> bool:
    return run_review(
        nda_path=nda_path,
        playbook_path=playbook_path,
        out_path=out_path,
        author=author,
        strict_ops=strict_ops,
        log=log,
    )


@app.post("/run")
async def run(
    nda: UploadFile = File(..., description="NDA .docx"),
    playbook: UploadFile = File(..., description="Playbook .xlsx"),
    author: str = Form(""),
    strict_ops: str = Form(""),
    verbose: str = Form(""),
) -> Response:
    """Flux sans SSE (compatibilité)."""
    if not nda.filename or not nda.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Le NDA doit être un fichier .docx")
    if not playbook.filename or not playbook.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Le playbook doit être un fichier .xlsx")

    log_buffer = io.StringIO()
    log = logging.getLogger(f"nda_web.{uuid.uuid4().hex}")
    log.handlers.clear()
    log.setLevel(logging.DEBUG if verbose == "1" else logging.INFO)
    log.propagate = False
    handler = logging.StreamHandler(log_buffer)
    handler.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    log.addHandler(handler)

    try:
        with tempfile.TemporaryDirectory(prefix="nda_web_") as td:
            tdir = Path(td)
            nda_path = tdir / "nda.docx"
            pb_path = tdir / "playbook.xlsx"
            out_path = tdir / "out.docx"

            content = await nda.read()
            if not content:
                raise HTTPException(status_code=400, detail="Fichier NDA vide")
            nda_path.write_bytes(content)

            pb_bytes = await playbook.read()
            if not pb_bytes:
                raise HTTPException(status_code=400, detail="Fichier playbook vide")
            pb_path.write_bytes(pb_bytes)

            strict = strict_ops == "1"
            ok = await run_in_threadpool(
                _run_review_sync,
                nda_path,
                pb_path,
                out_path,
                author,
                strict,
                log,
            )

            log_text = log_buffer.getvalue()
            if not ok:
                return JSONResponse(
                    status_code=500,
                    content={"detail": "La revue a échoué ou s’est arrêtée avant la sauvegarde.", "log": log_text},
                )

            if not out_path.is_file():
                return JSONResponse(
                    status_code=500,
                    content={"detail": "Fichier de sortie introuvable après traitement.", "log": log_text},
                )

            data = out_path.read_bytes()
    except HTTPException:
        raise
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"detail": str(e), "log": log_buffer.getvalue()},
        )

    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="NDA_revu.docx"'},
    )


def serve() -> None:
    import uvicorn

    uvicorn.run("nda_generator.web:app", host="127.0.0.1", port=8765, reload=False)


if __name__ == "__main__":
    serve()
