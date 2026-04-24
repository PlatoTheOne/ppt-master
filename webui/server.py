from __future__ import annotations

import json
import mimetypes
import os
import subprocess
import threading
import uuid
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import unquote, urlparse

from webui.pipeline import (
    REPO_ROOT,
    GenerationError,
    apply_draft_edits,
    create_draft,
    draft_state_path,
    export_draft,
    load_draft_state,
    retry_step,
)


STATIC_DIR = Path(__file__).resolve().parent / "static"
HOST = "127.0.0.1"
PORT = 8765

_jobs: dict[str, dict] = {}
_jobs_lock = threading.Lock()

STAGE_PRESETS = {
    "draft": {
        "queued": ("Queued", 5),
        "planning": ("Planning", 20),
        "rendering": ("Rendering draft", 55),
        "validating": ("Validating SVG", 80),
        "completed": ("Draft ready", 100),
    },
    "export": {
        "queued": ("Queued", 5),
        "editing": ("Applying edits", 20),
        "finalizing": ("Preparing outputs", 55),
        "exporting": ("Writing exports", 82),
        "completed": ("Export complete", 100),
    },
    "retry-render": {
        "queued": ("Queued", 5),
        "rendering": ("Re-rendering draft", 50),
        "validating": ("Validating SVG", 82),
        "completed": ("Draft refreshed", 100),
    },
    "retry-export": {
        "queued": ("Queued", 5),
        "finalizing": ("Preparing outputs", 50),
        "exporting": ("Writing exports", 82),
        "completed": ("Export complete", 100),
    },
}


def repo_relative(path: Path) -> str:
    return path.resolve().relative_to(REPO_ROOT.resolve()).as_posix()


def within_repo(path: Path) -> bool:
    try:
        path.resolve().relative_to(REPO_ROOT.resolve())
        return True
    except ValueError:
        return False


def decode_project_ref(project_ref: str) -> Path:
    target = (REPO_ROOT / unquote(project_ref)).resolve()
    if not within_repo(target):
        raise GenerationError("Project path is outside the repository.")
    return target


def plan_to_editor(plan: dict) -> dict:
    slides = []
    for slide in plan.get("slides") or []:
        slides.append(
            {
                "id": slide.get("id"),
                "type": slide.get("type"),
                "title": slide.get("title", ""),
                "subtitle": slide.get("subtitle", ""),
            }
        )
    return {
        "deckTitle": plan.get("deck_title", ""),
        "deckSubtitle": plan.get("deck_subtitle", ""),
        "slides": slides,
    }


def serialize_draft(project_dir: Path, plan: dict, slide_paths: list[Path], theme: dict, source_meta: dict | None = None) -> dict:
    return {
        "projectDir": repo_relative(project_dir),
        "draftState": repo_relative(draft_state_path(project_dir)),
        "slides": [f"/generated/{repo_relative(path)}" for path in slide_paths],
        "plan": plan,
        "editor": plan_to_editor(plan),
        "theme": theme,
        "sourceMeta": source_meta or {},
    }


def serialize_export(
    project_dir: Path,
    plan: dict,
    native: Path | None,
    legacy: Path | None,
    web_deck: Path | None,
    slide_paths: list[Path],
) -> dict:
    payload = {
        "projectDir": repo_relative(project_dir),
        "nativePptx": f"/generated/{repo_relative(native)}" if native else None,
        "legacyPptx": f"/generated/{repo_relative(legacy)}" if legacy else None,
        "webDeck": f"/generated/{repo_relative(web_deck)}" if web_deck else None,
        "slides": [f"/generated/{repo_relative(path)}" for path in slide_paths],
        "plan": plan,
        "editor": plan_to_editor(plan),
    }
    return payload


def list_history() -> list[dict]:
    projects_dir = REPO_ROOT / "projects"
    if not projects_dir.exists():
        return []

    history: list[dict] = []
    for project_dir in projects_dir.iterdir():
        if not project_dir.is_dir():
            continue
        state_file = draft_state_path(project_dir)
        exports_dir = project_dir / "exports"
        slide_dir = project_dir / "svg_final"
        if not slide_dir.exists():
            slide_dir = project_dir / "svg_output"
        slides = sorted(slide_dir.glob("*.svg"))[:8] if slide_dir.exists() else []

        item = {
            "projectDir": repo_relative(project_dir),
            "exportsDir": repo_relative(exports_dir) if exports_dir.exists() else None,
            "slides": [f"/generated/{repo_relative(path)}" for path in slides],
            "title": project_dir.name,
            "updatedAt": project_dir.stat().st_mtime,
            "hasDraft": state_file.exists(),
            "nativePptx": None,
            "legacyPptx": None,
            "webDeck": None,
        }

        if exports_dir.exists():
            native_candidates = sorted(
                [path for path in exports_dir.glob("*.pptx") if not path.name.endswith("_svg.pptx")],
                key=lambda path: path.stat().st_mtime,
                reverse=True,
            )
            if native_candidates:
                native = native_candidates[0]
                legacy = exports_dir / f"{native.stem}_svg.pptx"
                item["nativePptx"] = f"/generated/{repo_relative(native)}"
                item["legacyPptx"] = f"/generated/{repo_relative(legacy)}" if legacy.exists() else None
                item["updatedAt"] = native.stat().st_mtime
            web_candidates = sorted(exports_dir.glob("*_web_deck.html"), key=lambda path: path.stat().st_mtime, reverse=True)
            if web_candidates:
                web_deck = web_candidates[0]
                item["webDeck"] = f"/generated/{repo_relative(web_deck)}"
                item["updatedAt"] = max(item["updatedAt"], web_deck.stat().st_mtime)

        history.append(item)

    history.sort(key=lambda item: item["updatedAt"], reverse=True)
    return history[:12]


def open_repo_path(relative_path: str) -> None:
    target = decode_project_ref(relative_path)
    if not target.exists():
        raise FileNotFoundError(str(target))
    if target.is_file():
        subprocess.Popen(["explorer.exe", "/select,", str(target)])
        return
    subprocess.Popen(["explorer.exe", str(target)])


def set_job_stage(job_id: str, stage_key: str) -> None:
    with _jobs_lock:
        job = _jobs[job_id]
        preset = STAGE_PRESETS[job["action"]]
        label, progress = preset[stage_key]
        job["stage"] = stage_key
        job["stageLabel"] = label
        job["progress"] = progress


def update_job(job_id: str, **changes) -> None:
    with _jobs_lock:
        _jobs[job_id].update(changes)


def append_log(job_id: str, message: str) -> None:
    with _jobs_lock:
        _jobs[job_id]["logs"].append(message)


def create_job_record(action: str) -> str:
    job_id = uuid.uuid4().hex
    preset = STAGE_PRESETS[action]["queued"]
    with _jobs_lock:
        _jobs[job_id] = {
            "id": job_id,
            "action": action,
            "status": "queued",
            "stage": "queued",
            "stageLabel": preset[0],
            "progress": preset[1],
            "logs": [],
            "result": None,
            "error": None,
        }
    return job_id


def run_draft_job(job_id: str, payload: dict) -> None:
    update_job(job_id, status="running")
    try:
        set_job_stage(job_id, "planning")
        append_log(job_id, "Preparing draft outline")
        draft = create_draft(payload, lambda msg: append_log(job_id, msg))
        set_job_stage(job_id, "completed")
        result = serialize_draft(
            draft.project_dir,
            draft.plan,
            draft.slide_svgs,
            draft.theme,
            {"exportMode": payload.get("exportMode", "pptx")},
        )
        update_job(job_id, status="completed", result=result)
    except Exception as exc:
        message = str(exc) if isinstance(exc, GenerationError) else f"{exc.__class__.__name__}: {exc}"
        append_log(job_id, message)
        update_job(job_id, status="failed", error=message)


def run_export_job(job_id: str, payload: dict) -> None:
    update_job(job_id, status="running")
    try:
        project_dir = decode_project_ref(str(payload.get("projectDir") or ""))
        edits = payload.get("edits") or {}
        export_mode = str(payload.get("exportMode") or "pptx")
        if edits:
            set_job_stage(job_id, "editing")
            append_log(job_id, "Applying slide edits before export")
            draft = apply_draft_edits(project_dir, edits, lambda msg: append_log(job_id, msg))
            project_dir = draft.project_dir
        set_job_stage(job_id, "finalizing")
        append_log(job_id, "Preparing export artifacts")
        artifacts = export_draft(project_dir, lambda msg: append_log(job_id, msg), export_mode=export_mode)
        set_job_stage(job_id, "completed")
        result = serialize_export(
            artifacts.project_dir,
            artifacts.plan,
            artifacts.native_pptx,
            artifacts.legacy_pptx,
            artifacts.web_deck,
            artifacts.slide_svgs,
        )
        update_job(job_id, status="completed", result=result)
    except Exception as exc:
        message = str(exc) if isinstance(exc, GenerationError) else f"{exc.__class__.__name__}: {exc}"
        append_log(job_id, message)
        update_job(job_id, status="failed", error=message)


def run_retry_job(job_id: str, payload: dict) -> None:
    update_job(job_id, status="running")
    try:
        project_dir = decode_project_ref(str(payload.get("projectDir") or ""))
        step = str(payload.get("step") or "")
        export_mode = str(payload.get("exportMode") or "pptx")
        if step == "render":
            set_job_stage(job_id, "rendering")
        else:
            set_job_stage(job_id, "finalizing")
        result = retry_step(project_dir, step, lambda msg: append_log(job_id, msg), export_mode=export_mode if step == "export" else None)
        set_job_stage(job_id, "completed")

        if step == "render":
            draft_state = load_draft_state(project_dir)
            update_job(
                job_id,
                status="completed",
                result=serialize_draft(project_dir, draft_state["plan"], result["slides"], draft_state["theme"], draft_state.get("sourceMeta")),
            )
            return

        update_job(
            job_id,
            status="completed",
            result=serialize_export(
                result["projectDir"],
                result["plan"],
                result["nativePptx"],
                result["legacyPptx"],
                result.get("webDeck"),
                result["slides"],
            ),
        )
    except Exception as exc:
        message = str(exc) if isinstance(exc, GenerationError) else f"{exc.__class__.__name__}: {exc}"
        append_log(job_id, message)
        update_job(job_id, status="failed", error=message)


class WebUIHandler(BaseHTTPRequestHandler):
    server_version = "PPTMasterWebUI/0.2"

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/" or parsed.path == "/index.html":
            self.serve_file(STATIC_DIR / "index.html", "text/html; charset=utf-8")
            return
        if parsed.path.startswith("/static/"):
            rel = parsed.path.removeprefix("/static/")
            self.serve_file(STATIC_DIR / rel)
            return
        if parsed.path.startswith("/generated/"):
            rel = parsed.path.removeprefix("/generated/")
            file_path = REPO_ROOT / rel
            if not within_repo(file_path):
                self.send_error(HTTPStatus.FORBIDDEN, "Forbidden")
                return
            self.serve_file(file_path)
            return
        if parsed.path.startswith("/api/jobs/"):
            self.handle_job_status(parsed.path.rsplit("/", 1)[-1])
            return
        if parsed.path.startswith("/api/drafts/"):
            self.handle_get_draft(parsed.path.removeprefix("/api/drafts/"))
            return
        if parsed.path == "/api/config":
            self.send_json(
                {
                    "host": HOST,
                    "port": PORT,
                    "defaultModel": os.environ.get("OPENAI_MODEL", "gpt-4.1-mini"),
                }
            )
            return
        if parsed.path == "/api/history":
            self.send_json({"items": list_history()})
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/jobs":
            self.handle_create_draft_job()
            return
        if parsed.path == "/api/export-jobs":
            self.handle_create_export_job()
            return
        if parsed.path == "/api/retry-jobs":
            self.handle_create_retry_job()
            return
        if parsed.path == "/api/open-path":
            self.handle_open_path()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def read_json_body(self) -> dict | None:
        length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(length)
        try:
            return json.loads(raw_body.decode("utf-8"))
        except json.JSONDecodeError:
            self.send_json({"error": "Invalid JSON payload"}, status=HTTPStatus.BAD_REQUEST)
            return None

    def handle_create_draft_job(self) -> None:
        payload = self.read_json_body()
        if payload is None:
            return
        job_id = create_job_record("draft")
        worker = threading.Thread(target=run_draft_job, args=(job_id, payload), daemon=True)
        worker.start()
        self.send_json({"jobId": job_id}, status=HTTPStatus.ACCEPTED)

    def handle_create_export_job(self) -> None:
        payload = self.read_json_body()
        if payload is None:
            return
        job_id = create_job_record("export")
        worker = threading.Thread(target=run_export_job, args=(job_id, payload), daemon=True)
        worker.start()
        self.send_json({"jobId": job_id}, status=HTTPStatus.ACCEPTED)

    def handle_create_retry_job(self) -> None:
        payload = self.read_json_body()
        if payload is None:
            return
        step = str(payload.get("step") or "")
        action = "retry-render" if step == "render" else "retry-export"
        job_id = create_job_record(action)
        worker = threading.Thread(target=run_retry_job, args=(job_id, payload), daemon=True)
        worker.start()
        self.send_json({"jobId": job_id}, status=HTTPStatus.ACCEPTED)

    def handle_job_status(self, job_id: str) -> None:
        with _jobs_lock:
            job = _jobs.get(job_id)
            if not job:
                self.send_json({"error": "Job not found"}, status=HTTPStatus.NOT_FOUND)
                return
            self.send_json(job)

    def handle_get_draft(self, project_ref: str) -> None:
        try:
            project_dir = decode_project_ref(project_ref)
            state = load_draft_state(project_dir)
            slide_dir = project_dir / "svg_output"
            slides = sorted(slide_dir.glob("*.svg")) if slide_dir.exists() else []
            self.send_json(serialize_draft(project_dir, state["plan"], slides, state["theme"], state.get("sourceMeta")))
        except Exception as exc:
            message = str(exc) if isinstance(exc, GenerationError) else "Draft not found"
            self.send_json({"error": message}, status=HTTPStatus.NOT_FOUND)

    def handle_open_path(self) -> None:
        payload = self.read_json_body()
        if payload is None:
            return
        relative_path = str(payload.get("path") or "").strip()
        if not relative_path:
            self.send_json({"error": "Missing path"}, status=HTTPStatus.BAD_REQUEST)
            return
        try:
            open_repo_path(relative_path)
        except FileNotFoundError:
            self.send_json({"error": "Path not found"}, status=HTTPStatus.NOT_FOUND)
            return
        except ValueError as exc:
            self.send_json({"error": str(exc)}, status=HTTPStatus.BAD_REQUEST)
            return
        self.send_json({"ok": True})

    def serve_file(self, file_path: Path, content_type: str | None = None) -> None:
        if not file_path.exists() or not file_path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        guessed = content_type or mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        data = file_path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", guessed)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, payload: dict, status: HTTPStatus = HTTPStatus.OK) -> None:
        data = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, _format: str, *args) -> None:
        return


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), WebUIHandler)
    print(f"PPT Master Web UI running at http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
