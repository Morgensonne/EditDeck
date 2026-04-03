import base64
import binascii
import concurrent.futures
import hashlib
import json
import re
from datetime import datetime
from pathlib import Path
from threading import Lock, Thread
from typing import Any, Callable, Optional, Sequence
from uuid import uuid4

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from app.editable_ppt import EditableDeckPipeline
from app.pipeline import PPTImagePipeline, RuntimeConfig
from app.schemas import EditableDeckResult, SlideOutline, SlideResult
from app.settings import get_settings
from app.source_ingest import SourceDocumentProcessor, SourceFileInput

settings = get_settings()
pipeline = PPTImagePipeline(settings=settings)
editable_pipeline = EditableDeckPipeline(settings=settings)
source_processor = SourceDocumentProcessor(settings=settings)

BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"

app = FastAPI(
    title="One-Click PPT Generator",
    description="生成 PPT 图片、普通 PPT，并支持转换为整套可编辑 PPT。",
    version="0.5.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

Path(settings.output_root).mkdir(parents=True, exist_ok=True)
app.mount("/generated", StaticFiles(directory=settings.output_root), name="generated")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = Lock()
SESSIONS: dict[str, dict[str, Any]] = {}
SESSIONS_LOCK = Lock()
DATA_URL_RE = re.compile(r"^data:(?P<mime>[^;]+);base64,(?P<data>.+)$", re.DOTALL)


def _now_iso() -> str:
    return datetime.utcnow().isoformat(timespec="seconds") + "Z"


def _parse_slide_count(slide_count_raw: Optional[str]) -> Optional[int]:
    raw = (slide_count_raw or "").strip().lower()
    if raw in {"", "auto", "none", "null"}:
        return None
    try:
        value = int(raw)
    except ValueError as exc:
        raise ValueError("slide_count 必须是整数或 auto。") from exc
    if value < 1 or value > 20:
        raise ValueError("slide_count 范围必须在 1-20。")
    return value


def _parse_information_density(information_density_raw: Optional[str]) -> str:
    raw = (information_density_raw or "medium").strip().lower()
    if not raw:
        raw = "medium"
    if raw not in {"auto", "low", "medium", "high", "extra"}:
        raise ValueError("information_density 必须是 auto、low、medium、high 或 extra。")
    return raw


def _validate_style_inputs(
    style_description: Optional[str],
    style_bytes: Optional[bytes],
) -> None:
    if (style_description or "").strip() and style_bytes:
        raise ValueError("风格描述与风格模板图互斥，请二选一。")


def _decode_style_template_base64(style_template_base64: Optional[str]) -> tuple[Optional[bytes], Optional[str]]:
    raw = (style_template_base64 or "").strip()
    if not raw:
        return None, None

    match = DATA_URL_RE.fullmatch(raw)
    if not match:
        raise ValueError("style_template_base64 必须是 data:image/...;base64,... 格式。")

    mime = match.group("mime").strip().lower()
    if not mime.startswith("image/"):
        raise ValueError("style_template_base64 必须是图片 data URL。")

    try:
        data = base64.b64decode(match.group("data"), validate=True)
    except (ValueError, binascii.Error) as exc:
        raise ValueError("style_template_base64 不是有效的 base64 图片数据。") from exc

    if not data:
        raise ValueError("style_template_base64 不能为空图片。")

    return data, mime


async def _resolve_style_template_payload(
    style_template: Optional[UploadFile],
    style_template_base64: Optional[str],
) -> tuple[Optional[bytes], Optional[str]]:
    base64_bytes, base64_mime = _decode_style_template_base64(style_template_base64)
    if base64_bytes:
        return base64_bytes, base64_mime

    style_bytes = await style_template.read() if style_template else None
    style_mime = style_template.content_type if style_template else None
    return style_bytes, style_mime


def _parse_bool(value: Optional[str]) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def _job_snapshot(job_id: str) -> dict[str, Any]:
    with JOBS_LOCK:
        if job_id not in JOBS:
            raise KeyError(job_id)
        return dict(JOBS[job_id])


def _update_job(job_id: str, **updates: Any) -> None:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = _now_iso()


def _create_job(initial_message: str = "任务已创建，等待执行...") -> str:
    job_id = uuid4().hex
    now = _now_iso()
    with JOBS_LOCK:
        JOBS[job_id] = {
            "job_id": job_id,
            "state": "queued",
            "step": "queued",
            "message": initial_message,
            "progress": 0,
            "current_slide": 0,
            "total_slides": 0,
            "done": False,
            "error": "",
            "result": None,
            "result_preview": None,
            "created_at": now,
            "updated_at": now,
        }
    return job_id


def _scaled_progress_callback(
    target: Callable[[dict[str, Any]], None],
    start: int,
    end: int,
) -> Callable[[dict[str, Any]], None]:
    span = max(end - start, 1)

    def wrapped(payload: dict[str, Any]) -> None:
        raw_progress = int(payload.get("progress", 0) or 0)
        mapped = dict(payload)
        mapped["progress"] = min(end, start + int((max(0, min(100, raw_progress)) / 100) * span))
        target(mapped)

    return wrapped


def _resolve_run_dir(run_id: str) -> Path:
    run_dir = (Path(settings.output_root) / run_id).resolve()
    root_dir = Path(settings.output_root).resolve()
    try:
        run_dir.relative_to(root_dir)
    except ValueError as exc:
        raise ValueError("invalid run_id") from exc
    if not run_dir.exists():
        raise FileNotFoundError(f"run_id not found: {run_id}")
    return run_dir


def _build_editable_runtime_config(
    *,
    editable_base_url: Optional[str],
    editable_api_key: Optional[str],
    editable_model: Optional[str],
    editable_prompt_file: Optional[str],
    editable_browser_path: Optional[str],
    editable_download_timeout_ms: Optional[int],
    editable_max_tokens: Optional[int],
    editable_max_attempts: Optional[int],
    editable_sleep_seconds: Optional[float],
    assets_dir: Optional[str],
    asset_backend: Optional[str],
    mineru_base_url: Optional[str],
    mineru_api_key: Optional[str],
    mineru_model_version: Optional[str],
    mineru_language: Optional[str],
    mineru_enable_formula: Optional[bool],
    mineru_enable_table: Optional[bool],
    mineru_is_ocr: Optional[bool],
    mineru_poll_interval_seconds: Optional[float],
    mineru_timeout_seconds: Optional[int],
    mineru_max_refine_depth: Optional[int],
    force_reextract_assets: Optional[bool],
    disable_asset_reuse: Optional[bool],
):
    return editable_pipeline.build_runtime_config(
        base_url=editable_base_url,
        api_key=editable_api_key,
        model=editable_model,
        prompt_file=editable_prompt_file,
        chrome_path=editable_browser_path,
        download_timeout_ms=editable_download_timeout_ms,
        max_tokens=editable_max_tokens,
        max_attempts=editable_max_attempts,
        sleep_seconds=editable_sleep_seconds,
        assets_dir=assets_dir,
        asset_backend=asset_backend,
        mineru_base_url=mineru_base_url,
        mineru_api_key=mineru_api_key,
        mineru_model_version=mineru_model_version,
        mineru_language=mineru_language,
        mineru_enable_formula=mineru_enable_formula,
        mineru_enable_table=mineru_enable_table,
        mineru_is_ocr=mineru_is_ocr,
        mineru_poll_interval_seconds=mineru_poll_interval_seconds,
        mineru_timeout_seconds=mineru_timeout_seconds,
        mineru_max_refine_depth=mineru_max_refine_depth,
        force_reextract_assets=force_reextract_assets,
        disable_asset_reuse=disable_asset_reuse,
    )


def _build_source_runtime_config(
    *,
    text_provider: str,
    text_base_url: str,
    text_api_key: str,
    text_model: str,
    mineru_base_url: Optional[str],
    mineru_api_key: Optional[str],
    mineru_model_version: Optional[str],
    mineru_language: Optional[str],
    mineru_enable_formula: Optional[bool],
    mineru_enable_table: Optional[bool],
    mineru_is_ocr: Optional[bool],
    mineru_poll_interval_seconds: Optional[float],
    mineru_timeout_seconds: Optional[int],
):
    return source_processor.build_runtime_config(
        text_provider=text_provider,
        text_base_url=text_base_url,
        text_api_key=text_api_key,
        text_model=text_model,
        mineru_base_url=mineru_base_url,
        mineru_api_key=mineru_api_key,
        mineru_model_version=mineru_model_version,
        mineru_language=mineru_language,
        mineru_enable_formula=mineru_enable_formula,
        mineru_enable_table=mineru_enable_table,
        mineru_is_ocr=mineru_is_ocr,
        mineru_poll_interval_seconds=mineru_poll_interval_seconds,
        mineru_timeout_seconds=mineru_timeout_seconds,
    )


def _validate_editable_backend_args(
    *,
    asset_backend: Optional[str],
) -> None:
    backend = (asset_backend or "").strip().lower()
    if backend in {"", "edit", "mineru"}:
        return
    raise ValueError("editable backend only supports `edit`.")


def _run_generation_job(
    job_id: str,
    user_requirement: str,
    source_files: list[SourceFileInput],
    slide_count: Optional[int],
    information_density: str,
    style_description: Optional[str],
    style_bytes: Optional[bytes],
    style_mime: Optional[str],
    runtime_cfg,
    source_runtime_cfg,
    export_mode: str,
    generate_editable_ppt: bool,
    editable_runtime_cfg,
) -> None:
    _update_job(job_id, state="running", step="prepare", message="任务开始执行...", progress=1)

    def on_progress(payload: dict[str, Any]) -> None:
        _update_job(
            job_id,
            state="running",
            step=payload.get("step", "running"),
            message=payload.get("message", "处理中..."),
            progress=payload.get("progress", 0),
            current_slide=payload.get("current_slide", 0),
            total_slides=payload.get("total_slides", 0),
            done=payload.get("done", False),
            error=payload.get("error", ""),
        )

    try:
        prepared_requirement = source_processor.prepare_requirement(
            user_requirement=user_requirement,
            source_files=source_files,
            runtime_cfg=source_runtime_cfg,
        )
        generation_progress = on_progress
        editable_progress = on_progress
        if generate_editable_ppt:
            generation_progress = _scaled_progress_callback(on_progress, 0, 60)
            editable_progress = _scaled_progress_callback(on_progress, 60, 100)

        result = pipeline.run(
            user_requirement=prepared_requirement.final_requirement,
            slide_count=slide_count,
            style_description=style_description,
            style_template_bytes=style_bytes,
            style_template_mime=style_mime,
            runtime_cfg=runtime_cfg,
            export_mode=export_mode,
            information_density=information_density,
            progress_callback=generation_progress,
        )

        if generate_editable_ppt:
            editable_result = editable_pipeline.run_from_images(
                slide_images=[Path(slide.image_path) for slide in result.slides],
                runtime_cfg=editable_runtime_cfg,
                output_dir=Path(result.output_dir) / "editable_deck",
                progress_callback=editable_progress,
            )
            result = result.model_copy(update={"editable_deck": editable_result})

        _update_job(
            job_id,
            state="done",
            step="completed",
            message="生成完成",
            progress=100,
            done=True,
            error="",
            result=result.model_dump(),
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"生成失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


def _run_editable_job(job_id: str, run_id: str, editable_runtime_cfg) -> None:
    _update_job(job_id, state="running", step="editable_prepare", message="开始转换可编辑 PPT...", progress=1)

    def on_progress(payload: dict[str, Any]) -> None:
        _update_job(
            job_id,
            state="running",
            step=payload.get("step", "running"),
            message=payload.get("message", "处理中..."),
            progress=payload.get("progress", 0),
            current_slide=payload.get("current_slide", 0),
            total_slides=payload.get("total_slides", 0),
            done=payload.get("done", False),
            error=payload.get("error", ""),
        )

    try:
        run_dir = _resolve_run_dir(run_id)
        result = editable_pipeline.run_from_run_dir(
            run_dir=run_dir,
            runtime_cfg=editable_runtime_cfg,
            output_dir=run_dir / "editable_deck",
            progress_callback=on_progress,
        )
        _update_job(
            job_id,
            state="done",
            step="completed",
            message="可编辑 PPT 生成完成",
            progress=100,
            done=True,
            error="",
            result=result.model_dump(),
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"可编辑 PPT 生成失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


@app.get("/", include_in_schema=False)
def index() -> FileResponse:
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/generate")
async def generate_sync(
    user_requirement: str = Form(...),
    slide_count: Optional[str] = Form(default="auto"),
    information_density: Optional[str] = Form(default="medium"),
    style_description: Optional[str] = Form(default=None),
    style_template: Optional[UploadFile] = File(default=None),
    style_template_base64: Optional[str] = Form(default=None),
    source_files: Optional[list[UploadFile]] = File(default=None),
    export_mode: str = Form(default="both"),
    generate_editable_ppt: Optional[str] = Form(default="false"),
    base_url: Optional[str] = Form(default=None),
    image_api_url: Optional[str] = Form(default=None),
    text_api_key: Optional[str] = Form(default=None),
    image_api_key: Optional[str] = Form(default=None),
    text_model: Optional[str] = Form(default=None),
    image_model: Optional[str] = Form(default=None),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        resolved_slide_count = _parse_slide_count(slide_count)
        resolved_information_density = _parse_information_density(information_density)
        runtime_cfg = pipeline.build_runtime_config(
            base_url=base_url,
            image_api_url=image_api_url,
            text_api_key=text_api_key,
            image_api_key=image_api_key,
            text_model=text_model,
            image_model=image_model,
        )
        source_runtime_cfg = _build_source_runtime_config(
            text_provider=runtime_cfg.text_provider,
            text_base_url=runtime_cfg.text_base_url,
            text_api_key=runtime_cfg.text_api_key,
            text_model=runtime_cfg.text_model,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
        )
        editable_runtime_cfg = None
        if _parse_bool(generate_editable_ppt):
            _validate_editable_backend_args(
                asset_backend=asset_backend,
            )
            editable_runtime_cfg = _build_editable_runtime_config(
                editable_base_url=editable_base_url,
                editable_api_key=editable_api_key,
                editable_model=editable_model,
                editable_prompt_file=editable_prompt_file,
                editable_browser_path=editable_browser_path,
                editable_download_timeout_ms=editable_download_timeout_ms,
                editable_max_tokens=editable_max_tokens,
                editable_max_attempts=editable_max_attempts,
                editable_sleep_seconds=editable_sleep_seconds,
                assets_dir=assets_dir,
                asset_backend=asset_backend,
                mineru_base_url=mineru_base_url,
                mineru_api_key=mineru_api_key,
                mineru_model_version=mineru_model_version,
                mineru_language=mineru_language,
                mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
                mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
                mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
                mineru_poll_interval_seconds=mineru_poll_interval_seconds,
                mineru_timeout_seconds=mineru_timeout_seconds,
                mineru_max_refine_depth=mineru_max_refine_depth,
                force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
                disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
            )
        style_bytes, style_mime = await _resolve_style_template_payload(style_template, style_template_base64)
        source_payloads = [
            SourceFileInput(name=upload.filename or "source", data=await upload.read())
            for upload in (source_files or [])
            if upload and (upload.filename or "").strip()
        ]
        _validate_style_inputs(style_description=style_description, style_bytes=style_bytes)
        prepared_requirement = source_processor.prepare_requirement(
            user_requirement=user_requirement,
            source_files=source_payloads,
            runtime_cfg=source_runtime_cfg,
        )

        result = pipeline.run(
            user_requirement=prepared_requirement.final_requirement,
            slide_count=resolved_slide_count,
            style_description=style_description,
            style_template_bytes=style_bytes,
            style_template_mime=style_mime,
            runtime_cfg=runtime_cfg,
            export_mode=export_mode,
            information_density=resolved_information_density,
            progress_callback=None,
        )
        if _parse_bool(generate_editable_ppt):
            editable_result = editable_pipeline.run_from_images(
                slide_images=[Path(slide.image_path) for slide in result.slides],
                runtime_cfg=editable_runtime_cfg,
                output_dir=Path(result.output_dir) / "editable_deck",
                progress_callback=None,
            )
            result = result.model_copy(update={"editable_deck": editable_result})
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"生成失败：{exc}") from exc
    return result.model_dump()


@app.post("/api/generate/start")
async def generate_start(
    user_requirement: str = Form(...),
    slide_count: Optional[str] = Form(default="auto"),
    information_density: Optional[str] = Form(default="medium"),
    style_description: Optional[str] = Form(default=None),
    style_template: Optional[UploadFile] = File(default=None),
    style_template_base64: Optional[str] = Form(default=None),
    source_files: Optional[list[UploadFile]] = File(default=None),
    export_mode: str = Form(default="both"),
    generate_editable_ppt: Optional[str] = Form(default="false"),
    base_url: Optional[str] = Form(default=None),
    image_api_url: Optional[str] = Form(default=None),
    text_api_key: Optional[str] = Form(default=None),
    image_api_key: Optional[str] = Form(default=None),
    text_model: Optional[str] = Form(default=None),
    image_model: Optional[str] = Form(default=None),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        resolved_slide_count = _parse_slide_count(slide_count)
        resolved_information_density = _parse_information_density(information_density)
        runtime_cfg = pipeline.build_runtime_config(
            base_url=base_url,
            image_api_url=image_api_url,
            text_api_key=text_api_key,
            image_api_key=image_api_key,
            text_model=text_model,
            image_model=image_model,
        )
        source_runtime_cfg = _build_source_runtime_config(
            text_provider=runtime_cfg.text_provider,
            text_base_url=runtime_cfg.text_base_url,
            text_api_key=runtime_cfg.text_api_key,
            text_model=runtime_cfg.text_model,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
        )
        editable_runtime_cfg = None
        if _parse_bool(generate_editable_ppt):
            _validate_editable_backend_args(
                asset_backend=asset_backend,
            )
            editable_runtime_cfg = _build_editable_runtime_config(
                editable_base_url=editable_base_url,
                editable_api_key=editable_api_key,
                editable_model=editable_model,
                editable_prompt_file=editable_prompt_file,
                editable_browser_path=editable_browser_path,
                editable_download_timeout_ms=editable_download_timeout_ms,
                editable_max_tokens=editable_max_tokens,
                editable_max_attempts=editable_max_attempts,
                editable_sleep_seconds=editable_sleep_seconds,
                assets_dir=assets_dir,
                asset_backend=asset_backend,
                mineru_base_url=mineru_base_url,
                mineru_api_key=mineru_api_key,
                mineru_model_version=mineru_model_version,
                mineru_language=mineru_language,
                mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
                mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
                mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
                mineru_poll_interval_seconds=mineru_poll_interval_seconds,
                mineru_timeout_seconds=mineru_timeout_seconds,
                mineru_max_refine_depth=mineru_max_refine_depth,
                force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
                disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
            )
        style_bytes, style_mime = await _resolve_style_template_payload(style_template, style_template_base64)
        source_payloads = [
            SourceFileInput(name=upload.filename or "source", data=await upload.read())
            for upload in (source_files or [])
            if upload and (upload.filename or "").strip()
        ]
        _validate_style_inputs(style_description=style_description, style_bytes=style_bytes)

        job_id = _create_job()
        thread = Thread(
            target=_run_generation_job,
            args=(
                job_id,
                user_requirement,
                source_payloads,
                resolved_slide_count,
                resolved_information_density,
                style_description,
                style_bytes,
                style_mime,
                runtime_cfg,
                source_runtime_cfg,
                export_mode,
                _parse_bool(generate_editable_ppt),
                editable_runtime_cfg,
            ),
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued"}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"任务启动失败：{exc}") from exc


@app.post("/api/editable/start")
async def editable_start(
    run_id: str = Form(...),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        _validate_editable_backend_args(
            asset_backend=asset_backend,
        )
        editable_runtime_cfg = _build_editable_runtime_config(
            editable_base_url=editable_base_url,
            editable_api_key=editable_api_key,
            editable_model=editable_model,
            editable_prompt_file=editable_prompt_file,
            editable_browser_path=editable_browser_path,
            editable_download_timeout_ms=editable_download_timeout_ms,
            editable_max_tokens=editable_max_tokens,
            editable_max_attempts=editable_max_attempts,
            editable_sleep_seconds=editable_sleep_seconds,
            assets_dir=assets_dir,
            asset_backend=asset_backend,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
            mineru_max_refine_depth=mineru_max_refine_depth,
            force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
            disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
        )
        _resolve_run_dir(run_id)
        job_id = _create_job(initial_message="可编辑 PPT 转换任务已创建，等待执行...")
        thread = Thread(
            target=_run_editable_job,
            args=(job_id, run_id, editable_runtime_cfg),
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued"}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except FileNotFoundError as exc:
        raise HTTPException(status_code=404, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"任务启动失败：{exc}") from exc


@app.get("/api/generate/status/{job_id}")
def generate_status(job_id: str) -> dict[str, Any]:
    try:
        return _job_snapshot(job_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="job not found") from exc


SUPPORTED_REPLICA_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}


def _new_run_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid4().hex[:8]


def _parse_export_mode(export_mode_raw: Optional[str]) -> str:
    mode = (export_mode_raw or "both").strip().lower()
    if mode not in {"images", "ppt", "both"}:
        raise ValueError("export_mode must be one of: images, ppt, both.")
    return mode


def _path_to_generated_url(path_raw: Optional[str]) -> str:
    raw = (path_raw or "").strip()
    if not raw:
        return ""
    try:
        file_path = Path(raw).resolve()
        root_dir = Path(settings.output_root).resolve()
        rel_path = file_path.relative_to(root_dir)
    except Exception:
        return ""
    return f"/generated/{rel_path.as_posix()}"


def _ensure_run_dir(run_id: str) -> Path:
    run_dir = (Path(settings.output_root) / run_id).resolve()
    root_dir = Path(settings.output_root).resolve()
    try:
        run_dir.relative_to(root_dir)
    except ValueError as exc:
        raise ValueError("invalid run_id") from exc
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def _serialize_slide_payloads(slides: Sequence[dict[str, Any]]) -> list[dict[str, Any]]:
    out: list[dict[str, Any]] = []
    for item in slides:
        data = dict(item)
        page = int(data.get("page", 0) or 0)
        if page <= 0:
            continue
        if not data.get("image_url") and data.get("image_path"):
            data["image_url"] = _path_to_generated_url(str(data.get("image_path", "")))
        out.append(data)
    out.sort(key=lambda x: int(x.get("page", 0)))
    return out


def _serialize_editable_result(raw_result: EditableDeckResult) -> dict[str, Any]:
    payload = raw_result.model_dump() if isinstance(raw_result, EditableDeckResult) else dict(raw_result)
    payload["pptx_url"] = payload.get("pptx_url") or _path_to_generated_url(str(payload.get("pptx_path", "")))
    slides: list[dict[str, Any]] = []
    for item in payload.get("slides", []):
        slide = dict(item)
        slide["image_url"] = _path_to_generated_url(str(slide.get("image_path", "")))
        slide["assets_json_url"] = _path_to_generated_url(str(slide.get("assets_json_path", "")))
        slide["builder_url"] = _path_to_generated_url(str(slide.get("builder_path", "")))
        slide["preview_html_url"] = _path_to_generated_url(str(slide.get("preview_html_path", "")))
        slide["preview_pptx_url"] = _path_to_generated_url(str(slide.get("preview_pptx_path", "")))
        slides.append(slide)
    payload["slides"] = slides
    return payload


def _session_snapshot(session_id: str) -> dict[str, Any]:
    with SESSIONS_LOCK:
        if session_id not in SESSIONS:
            raise KeyError(session_id)
        session = dict(SESSIONS[session_id])
    session["outline"] = [dict(item) for item in session.get("outline", [])]
    session["slides"] = [dict(item) for item in session.get("slides", [])]
    editable_deck = session.get("editable_deck")
    if isinstance(editable_deck, dict):
        session["editable_deck"] = dict(editable_deck)
    return session


def _update_session(session_id: str, **updates: Any) -> None:
    with SESSIONS_LOCK:
        session = SESSIONS.get(session_id)
        if not session:
            return
        session.update(updates)
        session["updated_at"] = _now_iso()


def _create_session(payload: dict[str, Any]) -> str:
    session_id = uuid4().hex
    now = _now_iso()
    data = dict(payload)
    data.update(
        {
            "session_id": session_id,
            "created_at": now,
            "updated_at": now,
        }
    )
    with SESSIONS_LOCK:
        SESSIONS[session_id] = data
    return session_id


def _session_public_payload(session: dict[str, Any]) -> dict[str, Any]:
    return {
        "session_id": session["session_id"],
        "mode": session.get("mode", "generate"),
        "user_requirement": session.get("user_requirement", ""),
        "prepared_requirement": session.get("prepared_requirement", ""),
        "deck_title": session.get("deck_title", ""),
        "style_prompt": session.get("style_prompt", ""),
        "information_density": session.get("information_density", "medium"),
        "export_mode": session.get("export_mode", "both"),
        "run_id": session.get("run_id", ""),
        "output_dir": session.get("output_dir", ""),
        "pptx_url": session.get("pptx_url", ""),
        "pptx_path": session.get("pptx_path", ""),
        "outline": [dict(item) for item in session.get("outline", [])],
        "slides": _serialize_slide_payloads(session.get("slides", [])),
        "editable_deck": dict(session.get("editable_deck", {}) or {}),
        "source_files": [dict(item) for item in session.get("source_files", [])],
        "created_at": session.get("created_at", ""),
        "updated_at": session.get("updated_at", ""),
    }


def _parse_pages(pages_raw: Optional[str]) -> list[int]:
    raw = (pages_raw or "").strip()
    if not raw:
        return []
    values: list[int] = []
    if raw.startswith("["):
        payload = json.loads(raw)
        if not isinstance(payload, list):
            raise ValueError("pages must be a list of integers.")
        values = [int(x) for x in payload]
    else:
        cleaned = raw.replace("，", ",")
        values = [int(part.strip()) for part in cleaned.split(",") if part.strip()]
    unique = sorted({value for value in values if value > 0})
    return unique


def _parse_key_points(raw_points: Any) -> list[str]:
    if isinstance(raw_points, str):
        text = raw_points.replace("；", "\n").replace(";", "\n")
        return [line.strip() for line in text.splitlines() if line.strip()]
    if isinstance(raw_points, (list, tuple, set)):
        return [str(item).strip() for item in raw_points if str(item).strip()]
    return []


def _parse_outline_json(outline_json: str, information_density: str) -> list[SlideOutline]:
    try:
        payload = json.loads(outline_json)
    except Exception as exc:
        raise ValueError("outline_json must be valid JSON.") from exc
    slides_raw = payload.get("slides") if isinstance(payload, dict) else payload
    if not isinstance(slides_raw, list) or not slides_raw:
        raise ValueError("outline_json must contain a non-empty slides array.")
    if len(slides_raw) > 20:
        raise ValueError("outline cannot exceed 20 slides.")
    normalized_density = _parse_information_density(information_density)
    slides: list[SlideOutline] = []
    for index, raw_slide in enumerate(slides_raw, start=1):
        item = raw_slide if isinstance(raw_slide, dict) else {}
        title = str(item.get("title") or f"第{index}页").strip()
        if not title:
            title = f"第{index}页"
        key_points = pipeline._normalize_outline_key_points(
            _parse_key_points(item.get("key_points")),
            normalized_density,
        )
        slides.append(
            SlideOutline(
                page=index,
                title=title,
                key_points=key_points,
            )
        )
    return slides


def _resolve_runtime_bundle(
    *,
    base_url: Optional[str],
    image_api_url: Optional[str],
    text_api_key: Optional[str],
    image_api_key: Optional[str],
    text_model: Optional[str],
    image_model: Optional[str],
    mineru_base_url: Optional[str],
    mineru_api_key: Optional[str],
    mineru_model_version: Optional[str],
    mineru_language: Optional[str],
    mineru_enable_formula: Optional[str],
    mineru_enable_table: Optional[str],
    mineru_is_ocr: Optional[str],
    mineru_poll_interval_seconds: Optional[float],
    mineru_timeout_seconds: Optional[int],
) -> tuple[RuntimeConfig, Any]:
    runtime_cfg = pipeline.build_runtime_config(
        base_url=base_url,
        image_api_url=image_api_url,
        text_api_key=text_api_key,
        image_api_key=image_api_key,
        text_model=text_model,
        image_model=image_model,
    )
    source_runtime_cfg = _build_source_runtime_config(
        text_provider=runtime_cfg.text_provider,
        text_base_url=runtime_cfg.text_base_url,
        text_api_key=runtime_cfg.text_api_key,
        text_model=runtime_cfg.text_model,
        mineru_base_url=mineru_base_url,
        mineru_api_key=mineru_api_key,
        mineru_model_version=mineru_model_version,
        mineru_language=mineru_language,
        mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
        mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
        mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
        mineru_poll_interval_seconds=mineru_poll_interval_seconds,
        mineru_timeout_seconds=mineru_timeout_seconds,
    )
    return runtime_cfg, source_runtime_cfg


def _build_result_payload_from_session(session_id: str) -> dict[str, Any]:
    return _session_public_payload(_session_snapshot(session_id))


def _run_workflow_render_job(
    *,
    job_id: str,
    session_id: str,
    export_mode: str,
    selected_pages: list[int],
) -> None:
    _update_job(
        job_id,
        state="running",
        step="prepare",
        message="正在准备渲染任务...",
        progress=2,
        current_slide=0,
        total_slides=0,
        done=False,
        error="",
    )
    try:
        session = _session_snapshot(session_id)
        runtime_cfg: RuntimeConfig = session["runtime_cfg"]
        if runtime_cfg is None:
            raise ValueError("当前会话不支持文本生成渲染，请使用图片复刻模式。")
        requirement = str(session.get("prepared_requirement", "")).strip()
        deck_title = str(session.get("deck_title", "")).strip() or "自动生成PPT"
        style_prompt = str(session.get("style_prompt", "")).strip()
        information_density = _parse_information_density(session.get("information_density", "medium"))
        outline = [SlideOutline.model_validate(item) for item in session.get("outline", [])]
        if not outline:
            raise ValueError("当前会话没有可用大纲，请先完成步骤1。")
        page_map = {slide.page: slide for slide in outline}

        target_pages = selected_pages or [slide.page for slide in outline]
        for page in target_pages:
            if page not in page_map:
                raise ValueError(f"page {page} 不在当前大纲范围内。")

        run_id = str(session.get("run_id") or "").strip() or _new_run_id()
        run_dir = _ensure_run_dir(run_id)

        existing_slides = {
            int(item["page"]): dict(item)
            for item in session.get("slides", [])
            if int(item.get("page", 0) or 0) > 0
        }
        if not selected_pages:
            existing_slides = {}

        style_bytes: bytes | None = session.get("style_template_bytes")
        style_mime: str | None = session.get("style_template_mime")
        style_reference_data_url = pipeline._image_bytes_to_data_url(style_bytes, style_mime)
        style_reference_sha256 = hashlib.sha256(style_bytes).hexdigest() if style_bytes else ""
        style_reference_mime = (style_mime or "image/png") if style_reference_data_url else None

        total_pages = len(target_pages)
        _update_job(
            job_id,
            state="running",
            step="prompt_generation",
            message="正在生成逐页渲染 Prompt...",
            progress=8,
            current_slide=0,
            total_slides=total_pages,
            done=False,
            error="",
        )

        page_prompts: dict[int, str] = {}
        for index, page in enumerate(target_pages, start=1):
            slide = page_map[page]
            page_prompts[page] = pipeline._generate_slide_render_prompt(
                deck_title=deck_title,
                requirement=requirement,
                slide=slide,
                style_prompt=style_prompt,
                runtime_cfg=runtime_cfg,
                information_density=information_density,
                style_reference_data_url=style_reference_data_url,
                style_reference_mime=style_reference_mime,
                style_reference_sha256=style_reference_sha256,
                logger=None,
            )
            _update_job(
                job_id,
                state="running",
                step="prompt_generation",
                message=f"Prompt 已完成 {index}/{total_pages}",
                progress=8 + int((index / max(1, total_pages)) * 24),
                current_slide=index,
                total_slides=total_pages,
                done=False,
                error="",
            )

        workers = max(1, min(settings.image_max_workers, total_pages))
        _update_job(
            job_id,
            state="running",
            step="image_generation",
            message=f"正在并发生成图片（{workers} 线程）...",
            progress=35,
            current_slide=0,
            total_slides=total_pages,
            done=False,
            error="",
        )

        completed = 0
        with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
            futures: dict[concurrent.futures.Future[SlideResult], int] = {}
            for page in target_pages:
                slide = page_map[page]
                futures[
                    executor.submit(
                        pipeline._render_one_slide,
                        runtime_cfg,
                        run_id,
                        run_dir,
                        slide,
                        page_prompts[page],
                        None,
                    )
                ] = page
            try:
                for future in concurrent.futures.as_completed(futures):
                    slide_result = future.result()
                    slide_payload = slide_result.model_dump()
                    slide_payload["rendered_at"] = _now_iso()
                    existing_slides[slide_result.page] = slide_payload
                    completed += 1
                    ordered = _serialize_slide_payloads(list(existing_slides.values()))
                    _update_session(
                        session_id,
                        run_id=run_id,
                        output_dir=str(run_dir.resolve()),
                        slides=ordered,
                        editable_deck={},
                    )
                    preview = _build_result_payload_from_session(session_id)
                    _update_job(
                        job_id,
                        state="running",
                        step="image_generation",
                        message=f"页面已生成 {completed}/{total_pages}",
                        progress=35 + int((completed / max(1, total_pages)) * 55),
                        current_slide=completed,
                        total_slides=total_pages,
                        done=False,
                        error="",
                        result_preview=preview,
                    )
            except Exception:
                for pending in futures:
                    pending.cancel()
                raise

        ordered_slide_models: list[SlideResult] = []
        for slide in outline:
            payload = existing_slides.get(slide.page)
            if payload:
                ordered_slide_models.append(SlideResult.model_validate(payload))

        pptx_url = ""
        pptx_path = ""
        if export_mode in {"ppt", "both"} and len(ordered_slide_models) == len(outline):
            _update_job(
                job_id,
                state="running",
                step="packaging",
                message="正在打包PPT...",
                progress=94,
                current_slide=total_pages,
                total_slides=total_pages,
                done=False,
                error="",
            )
            pptx_name = "generated_deck.pptx"
            target_pptx_path = run_dir / pptx_name
            pipeline._build_pptx(ordered_slide_models, run_dir, target_pptx_path)
            pptx_path = str(target_pptx_path.resolve())
            pptx_url = f"/generated/{run_id}/{pptx_name}"

        session_updates: dict[str, Any] = {
            "run_id": run_id,
            "output_dir": str(run_dir.resolve()),
            "slides": _serialize_slide_payloads(list(existing_slides.values())),
            "export_mode": export_mode,
            "editable_deck": {},
        }
        if export_mode == "images":
            session_updates["pptx_url"] = ""
            session_updates["pptx_path"] = ""
        elif pptx_url:
            session_updates["pptx_url"] = pptx_url
            session_updates["pptx_path"] = pptx_path
        _update_session(session_id, **session_updates)
        final_result = _build_result_payload_from_session(session_id)
        _update_job(
            job_id,
            state="done",
            step="completed",
            message="渲染完成",
            progress=100,
            current_slide=total_pages,
            total_slides=total_pages,
            done=True,
            error="",
            result=final_result,
            result_preview=final_result,
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"渲染失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


def _run_workflow_editable_job(
    *,
    job_id: str,
    session_id: str,
    editable_runtime_cfg: Any,
    selected_pages: list[int],
) -> None:
    _update_job(
        job_id,
        state="running",
        step="editable_prepare",
        message="正在准备可编辑PPT任务...",
        progress=3,
        current_slide=0,
        total_slides=0,
        done=False,
        error="",
    )
    try:
        session = _session_snapshot(session_id)
        run_id = str(session.get("run_id", "")).strip()
        if not run_id:
            raise ValueError("当前会话尚未生成图片，无法转换可编辑PPT。")
        run_dir = _resolve_run_dir(run_id)

        def on_progress(payload: dict[str, Any]) -> None:
            preview = _build_result_payload_from_session(session_id)
            _update_job(
                job_id,
                state="running",
                step=payload.get("step", "running"),
                message=payload.get("message", "处理中..."),
                progress=payload.get("progress", 0),
                current_slide=payload.get("current_slide", 0),
                total_slides=payload.get("total_slides", 0),
                done=payload.get("done", False),
                error=payload.get("error", ""),
                result_preview=preview,
            )

        if selected_pages:
            slide_images = []
            for page in selected_pages:
                matches = sorted(run_dir.glob(f"slide_{page:02d}.*"))
                if not matches:
                    raise ValueError(f"未找到第 {page} 页图片，请先生成该页。")
                slide_images.append(matches[0])
            output_dir = run_dir / f"editable_partial_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            editable_result = editable_pipeline.run_from_images(
                slide_images=slide_images,
                runtime_cfg=editable_runtime_cfg,
                output_dir=output_dir,
                progress_callback=on_progress,
            )
            decorated_partial = _serialize_editable_result(editable_result)
            payload = _build_result_payload_from_session(session_id)
            payload["partial_editable_deck"] = decorated_partial
            _update_job(
                job_id,
                state="done",
                step="completed",
                message="选中页可编辑预览生成完成",
                progress=100,
                done=True,
                error="",
                result=payload,
                result_preview=payload,
            )
            return

        editable_result = editable_pipeline.run_from_run_dir(
            run_dir=run_dir,
            runtime_cfg=editable_runtime_cfg,
            output_dir=run_dir / "editable_deck",
            progress_callback=on_progress,
        )
        decorated = _serialize_editable_result(editable_result)
        _update_session(session_id, editable_deck=decorated)
        payload = _build_result_payload_from_session(session_id)
        _update_job(
            job_id,
            state="done",
            step="completed",
            message="可编辑PPT已完成",
            progress=100,
            done=True,
            error="",
            result=payload,
            result_preview=payload,
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"可编辑PPT失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


def _run_replica_job(
    *,
    job_id: str,
    session_id: str,
    replica_images: list[dict[str, Any]],
    export_mode: str,
    generate_editable_ppt: bool,
    editable_runtime_cfg: Any,
) -> None:
    _update_job(
        job_id,
        state="running",
        step="prepare",
        message="正在准备图片复刻任务...",
        progress=3,
        current_slide=0,
        total_slides=len(replica_images),
        done=False,
        error="",
    )
    try:
        _session_snapshot(session_id)
        run_id = _new_run_id()
        run_dir = _ensure_run_dir(run_id)

        slides: list[SlideResult] = []
        outline: list[dict[str, Any]] = []
        total = len(replica_images)
        for index, image in enumerate(replica_images, start=1):
            suffix = str(image.get("suffix") or ".png").lower()
            if suffix not in SUPPORTED_REPLICA_SUFFIXES:
                suffix = ".png"
            target_name = f"slide_{index:02d}{suffix}"
            target_path = run_dir / target_name
            target_path.write_bytes(image["data"])
            title = f"第{index}页"
            slide = SlideResult(
                page=index,
                title=title,
                prompt="图片复刻",
                image_url=f"/generated/{run_id}/{target_name}",
                image_path=str(target_path.resolve()),
            )
            slides.append(slide)
            outline.append({"page": index, "title": title, "key_points": [str(image.get("name") or title)]})
            slide_payloads = []
            for item in slides:
                payload = item.model_dump()
                payload["rendered_at"] = _now_iso()
                slide_payloads.append(payload)
            _update_session(
                session_id,
                run_id=run_id,
                output_dir=str(run_dir.resolve()),
                outline=outline,
                slides=slide_payloads,
                pptx_url="",
                pptx_path="",
                editable_deck={},
            )
            preview = _build_result_payload_from_session(session_id)
            _update_job(
                job_id,
                state="running",
                step="image_generation",
                message=f"已写入图片 {index}/{total}",
                progress=5 + int((index / max(1, total)) * 55),
                current_slide=index,
                total_slides=total,
                done=False,
                error="",
                result_preview=preview,
            )

        if export_mode in {"ppt", "both"}:
            _update_job(
                job_id,
                state="running",
                step="packaging",
                message="正在打包PPT...",
                progress=70,
                current_slide=total,
                total_slides=total,
                done=False,
                error="",
            )
            target_pptx = run_dir / "generated_deck.pptx"
            pipeline._build_pptx(slides, run_dir, target_pptx)
            _update_session(
                session_id,
                pptx_path=str(target_pptx.resolve()),
                pptx_url=f"/generated/{run_id}/generated_deck.pptx",
            )

        if generate_editable_ppt:
            def on_progress(payload: dict[str, Any]) -> None:
                _update_job(
                    job_id,
                    state="running",
                    step=payload.get("step", "running"),
                    message=payload.get("message", "处理中..."),
                    progress=70 + int((payload.get("progress", 0) / 100) * 30),
                    current_slide=payload.get("current_slide", 0),
                    total_slides=payload.get("total_slides", total),
                    done=False,
                    error=payload.get("error", ""),
                    result_preview=_build_result_payload_from_session(session_id),
                )

            editable_result = editable_pipeline.run_from_images(
                slide_images=[Path(item.image_path) for item in slides],
                runtime_cfg=editable_runtime_cfg,
                output_dir=run_dir / "editable_deck",
                progress_callback=on_progress,
            )
            _update_session(session_id, editable_deck=_serialize_editable_result(editable_result))

        final_payload = _build_result_payload_from_session(session_id)
        _update_job(
            job_id,
            state="done",
            step="completed",
            message="图片复刻完成",
            progress=100,
            current_slide=total,
            total_slides=total,
            done=True,
            error="",
            result=final_payload,
            result_preview=final_payload,
        )
    except Exception as exc:
        _update_job(
            job_id,
            state="failed",
            step="failed",
            message=f"图片复刻失败：{exc}",
            progress=100,
            done=True,
            error=str(exc),
            result=None,
        )


@app.get("/api/workflow/defaults")
def workflow_defaults() -> dict[str, Any]:
    return {
        "app": {
            "default_slide_count": settings.default_slide_count,
            "output_root": settings.output_root,
        },
        "models": {
            "text": {
                "provider": settings.text_provider,
                "base_url": settings.text_base_url,
                "model": settings.text_model,
                "has_api_key": bool(settings.text_api_key),
            },
            "image": {
                "provider": settings.image_provider,
                "base_url": settings.image_api_url,
                "model": settings.image_model,
                "size": settings.image_size,
                "has_api_key": bool(settings.image_api_key or settings.text_api_key),
            },
            "editable": {
                "provider": settings.editable_ppt_provider,
                "base_url": settings.resolved_editable_base_url,
                "model": settings.editable_ppt_model,
                "has_api_key": bool(settings.resolved_editable_api_key),
            },
        },
        "mineru": {
            "base_url": settings.resolved_mineru_base_url,
            "has_api_key": bool(settings.resolved_mineru_api_key),
            "model_version": settings.mineru_model_version,
            "language": settings.mineru_language,
            "enable_formula": settings.mineru_enable_formula,
            "enable_table": settings.mineru_enable_table,
            "is_ocr": settings.mineru_is_ocr,
        },
    }


@app.get("/api/workflow/session/{session_id}")
def workflow_session(session_id: str) -> dict[str, Any]:
    try:
        return _build_result_payload_from_session(session_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="session not found") from exc


@app.post("/api/workflow/prepare")
async def workflow_prepare(
    user_requirement: str = Form(...),
    slide_count: Optional[str] = Form(default="auto"),
    information_density: Optional[str] = Form(default="medium"),
    style_description: Optional[str] = Form(default=None),
    style_template: Optional[UploadFile] = File(default=None),
    style_template_base64: Optional[str] = Form(default=None),
    source_files: Optional[list[UploadFile]] = File(default=None),
    base_url: Optional[str] = Form(default=None),
    image_api_url: Optional[str] = Form(default=None),
    text_api_key: Optional[str] = Form(default=None),
    image_api_key: Optional[str] = Form(default=None),
    text_model: Optional[str] = Form(default=None),
    image_model: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
) -> dict[str, Any]:
    try:
        requested_slide_count = _parse_slide_count(slide_count)
        resolved_density = _parse_information_density(information_density)
        runtime_cfg, source_runtime_cfg = _resolve_runtime_bundle(
            base_url=base_url,
            image_api_url=image_api_url,
            text_api_key=text_api_key,
            image_api_key=image_api_key,
            text_model=text_model,
            image_model=image_model,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=mineru_enable_formula,
            mineru_enable_table=mineru_enable_table,
            mineru_is_ocr=mineru_is_ocr,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
        )
        style_bytes, style_mime = await _resolve_style_template_payload(style_template, style_template_base64)
        _validate_style_inputs(style_description=style_description, style_bytes=style_bytes)
        source_payloads = [
            SourceFileInput(name=upload.filename or "source", data=await upload.read())
            for upload in (source_files or [])
            if upload and (upload.filename or "").strip()
        ]
        prepared = source_processor.prepare_requirement(
            user_requirement=user_requirement,
            source_files=source_payloads,
            runtime_cfg=source_runtime_cfg,
        )
        resolved_slide_count = pipeline._resolve_slide_count(
            requirement=prepared.final_requirement,
            requested_slide_count=requested_slide_count,
            runtime_cfg=runtime_cfg,
            logger=None,
        )
        style_prompt = pipeline._generate_style_prompt(
            requirement=prepared.final_requirement,
            style_description=(style_description or "").strip(),
            style_template_bytes=style_bytes,
            style_template_mime=style_mime,
            runtime_cfg=runtime_cfg,
            logger=None,
        )
        outline_result = pipeline._generate_outline(
            requirement=prepared.final_requirement,
            slide_count=resolved_slide_count,
            information_density=resolved_density,
            runtime_cfg=runtime_cfg,
            logger=None,
        )
        session_id = _create_session(
            {
                "mode": "generate",
                "user_requirement": (user_requirement or "").strip(),
                "prepared_requirement": prepared.final_requirement,
                "source_files": [
                    {
                        "name": source.name,
                        "suffix": source.suffix,
                        "method": source.extraction_method,
                        "char_count": len(source.text),
                    }
                    for source in prepared.extracted_sources
                ],
                "runtime_cfg": runtime_cfg,
                "source_runtime_cfg": source_runtime_cfg,
                "style_template_bytes": style_bytes,
                "style_template_mime": style_mime,
                "deck_title": outline_result.deck_title,
                "style_prompt": style_prompt,
                "information_density": resolved_density,
                "outline": [slide.model_dump() for slide in outline_result.slides],
                "slides": [],
                "run_id": "",
                "output_dir": "",
                "pptx_url": "",
                "pptx_path": "",
                "editable_deck": {},
                "export_mode": "both",
            }
        )
        return _build_result_payload_from_session(session_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"prepare failed: {exc}") from exc


@app.post("/api/workflow/session/update")
def workflow_session_update(
    session_id: str = Form(...),
    deck_title: Optional[str] = Form(default=None),
    style_prompt: Optional[str] = Form(default=None),
    information_density: Optional[str] = Form(default=None),
    outline_json: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        session = _session_snapshot(session_id)
        updates: dict[str, Any] = {}
        if deck_title is not None:
            updates["deck_title"] = (deck_title or "").strip() or "自动生成PPT"
        if style_prompt is not None:
            cleaned_style_prompt = (style_prompt or "").strip()
            if not cleaned_style_prompt:
                raise ValueError("style_prompt cannot be empty.")
            updates["style_prompt"] = cleaned_style_prompt
        next_density = session.get("information_density", "medium")
        if information_density is not None:
            next_density = _parse_information_density(information_density)
            updates["information_density"] = next_density
        if outline_json:
            parsed_outline = _parse_outline_json(outline_json, next_density)
            updates["outline"] = [slide.model_dump() for slide in parsed_outline]
            updates["editable_deck"] = {}
        if updates:
            _update_session(session_id, **updates)
        return _build_result_payload_from_session(session_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="session not found") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"session update failed: {exc}") from exc


@app.post("/api/workflow/render/start")
def workflow_render_start(
    session_id: str = Form(...),
    export_mode: Optional[str] = Form(default="both"),
    render_pages: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        _session_snapshot(session_id)
        resolved_export_mode = _parse_export_mode(export_mode)
        pages = _parse_pages(render_pages)
        job_id = _create_job(initial_message="渲染任务已创建，等待执行...")
        thread = Thread(
            target=_run_workflow_render_job,
            kwargs={
                "job_id": job_id,
                "session_id": session_id,
                "export_mode": resolved_export_mode,
                "selected_pages": pages,
            },
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued", "session_id": session_id}
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="session not found") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"render start failed: {exc}") from exc


@app.post("/api/workflow/slides/regenerate/start")
def workflow_regenerate_start(
    session_id: str = Form(...),
    pages: str = Form(...),
    export_mode: Optional[str] = Form(default="both"),
) -> dict[str, Any]:
    try:
        resolved_pages = _parse_pages(pages)
        if not resolved_pages:
            raise ValueError("请至少指定一页进行重绘。")
        return workflow_render_start(session_id=session_id, export_mode=export_mode, render_pages=json.dumps(resolved_pages))
    except HTTPException:
        raise
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"regenerate start failed: {exc}") from exc


@app.post("/api/workflow/editable/start")
def workflow_editable_start(
    session_id: str = Form(...),
    editable_pages: Optional[str] = Form(default=None),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        _session_snapshot(session_id)
        _validate_editable_backend_args(asset_backend=asset_backend)
        editable_runtime_cfg = _build_editable_runtime_config(
            editable_base_url=editable_base_url,
            editable_api_key=editable_api_key,
            editable_model=editable_model,
            editable_prompt_file=editable_prompt_file,
            editable_browser_path=editable_browser_path,
            editable_download_timeout_ms=editable_download_timeout_ms,
            editable_max_tokens=editable_max_tokens,
            editable_max_attempts=editable_max_attempts,
            editable_sleep_seconds=editable_sleep_seconds,
            assets_dir=assets_dir,
            asset_backend=asset_backend,
            mineru_base_url=mineru_base_url,
            mineru_api_key=mineru_api_key,
            mineru_model_version=mineru_model_version,
            mineru_language=mineru_language,
            mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
            mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
            mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
            mineru_poll_interval_seconds=mineru_poll_interval_seconds,
            mineru_timeout_seconds=mineru_timeout_seconds,
            mineru_max_refine_depth=mineru_max_refine_depth,
            force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
            disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
        )
        selected_pages = _parse_pages(editable_pages)
        job_id = _create_job(initial_message="可编辑PPT任务已创建，等待执行...")
        thread = Thread(
            target=_run_workflow_editable_job,
            kwargs={
                "job_id": job_id,
                "session_id": session_id,
                "editable_runtime_cfg": editable_runtime_cfg,
                "selected_pages": selected_pages,
            },
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued", "session_id": session_id}
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="session not found") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"editable start failed: {exc}") from exc


@app.post("/api/workflow/replica/start")
async def workflow_replica_start(
    deck_title: Optional[str] = Form(default="图片复刻结果"),
    export_mode: Optional[str] = Form(default="both"),
    generate_editable_ppt: Optional[str] = Form(default="false"),
    replica_images: list[UploadFile] = File(...),
    editable_base_url: Optional[str] = Form(default=None),
    editable_api_key: Optional[str] = Form(default=None),
    editable_model: Optional[str] = Form(default=None),
    editable_prompt_file: Optional[str] = Form(default=None),
    editable_browser_path: Optional[str] = Form(default=None),
    editable_download_timeout_ms: Optional[int] = Form(default=None),
    editable_max_tokens: Optional[int] = Form(default=None),
    editable_max_attempts: Optional[int] = Form(default=None),
    editable_sleep_seconds: Optional[float] = Form(default=None),
    assets_dir: Optional[str] = Form(default=None),
    asset_backend: Optional[str] = Form(default=None),
    mineru_base_url: Optional[str] = Form(default=None),
    mineru_api_key: Optional[str] = Form(default=None),
    mineru_model_version: Optional[str] = Form(default=None),
    mineru_language: Optional[str] = Form(default=None),
    mineru_enable_formula: Optional[str] = Form(default=None),
    mineru_enable_table: Optional[str] = Form(default=None),
    mineru_is_ocr: Optional[str] = Form(default=None),
    mineru_poll_interval_seconds: Optional[float] = Form(default=None),
    mineru_timeout_seconds: Optional[int] = Form(default=None),
    mineru_max_refine_depth: Optional[int] = Form(default=None),
    force_reextract_assets: Optional[str] = Form(default=None),
    disable_asset_reuse: Optional[str] = Form(default=None),
) -> dict[str, Any]:
    try:
        if not replica_images:
            raise ValueError("请至少上传一张用于复刻的图片。")
        if len(replica_images) > 20:
            raise ValueError("复刻模式最多支持20张图片。")
        resolved_export_mode = _parse_export_mode(export_mode)
        replica_payloads: list[dict[str, Any]] = []
        for upload in replica_images:
            name = (upload.filename or "").strip() or "slide"
            suffix = Path(name).suffix.lower() or ".png"
            if suffix not in SUPPORTED_REPLICA_SUFFIXES:
                raise ValueError(f"不支持文件类型：{name}")
            replica_payloads.append(
                {
                    "name": name,
                    "suffix": suffix,
                    "data": await upload.read(),
                }
            )
        replica_payloads.sort(key=lambda item: item["name"].lower())

        should_generate_editable = _parse_bool(generate_editable_ppt)
        editable_runtime_cfg = None
        if should_generate_editable:
            _validate_editable_backend_args(asset_backend=asset_backend)
            editable_runtime_cfg = _build_editable_runtime_config(
                editable_base_url=editable_base_url,
                editable_api_key=editable_api_key,
                editable_model=editable_model,
                editable_prompt_file=editable_prompt_file,
                editable_browser_path=editable_browser_path,
                editable_download_timeout_ms=editable_download_timeout_ms,
                editable_max_tokens=editable_max_tokens,
                editable_max_attempts=editable_max_attempts,
                editable_sleep_seconds=editable_sleep_seconds,
                assets_dir=assets_dir,
                asset_backend=asset_backend,
                mineru_base_url=mineru_base_url,
                mineru_api_key=mineru_api_key,
                mineru_model_version=mineru_model_version,
                mineru_language=mineru_language,
                mineru_enable_formula=_parse_bool(mineru_enable_formula) if mineru_enable_formula is not None else None,
                mineru_enable_table=_parse_bool(mineru_enable_table) if mineru_enable_table is not None else None,
                mineru_is_ocr=_parse_bool(mineru_is_ocr) if mineru_is_ocr is not None else None,
                mineru_poll_interval_seconds=mineru_poll_interval_seconds,
                mineru_timeout_seconds=mineru_timeout_seconds,
                mineru_max_refine_depth=mineru_max_refine_depth,
                force_reextract_assets=_parse_bool(force_reextract_assets) if force_reextract_assets is not None else None,
                disable_asset_reuse=_parse_bool(disable_asset_reuse) if disable_asset_reuse is not None else None,
            )

        session_id = _create_session(
            {
                "mode": "replica",
                "user_requirement": "",
                "prepared_requirement": "",
                "source_files": [],
                "runtime_cfg": None,
                "source_runtime_cfg": None,
                "style_template_bytes": None,
                "style_template_mime": None,
                "deck_title": (deck_title or "").strip() or "图片复刻结果",
                "style_prompt": "",
                "information_density": "medium",
                "outline": [],
                "slides": [],
                "run_id": "",
                "output_dir": "",
                "pptx_url": "",
                "pptx_path": "",
                "editable_deck": {},
                "export_mode": resolved_export_mode,
            }
        )

        job_id = _create_job(initial_message="图片复刻任务已创建，等待执行...")
        thread = Thread(
            target=_run_replica_job,
            kwargs={
                "job_id": job_id,
                "session_id": session_id,
                "replica_images": replica_payloads,
                "export_mode": resolved_export_mode,
                "generate_editable_ppt": should_generate_editable,
                "editable_runtime_cfg": editable_runtime_cfg,
            },
            daemon=True,
        )
        thread.start()
        return {"job_id": job_id, "state": "queued", "session_id": session_id}
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"replica start failed: {exc}") from exc
