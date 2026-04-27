from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Any, Literal, Optional

import yaml
from pydantic import BaseModel, Field


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = PROJECT_ROOT / "config"
DEFAULT_CONFIG_FILE = CONFIG_DIR / "app.yaml"


class AppConfig(BaseModel):
    output_root: str = "generated"
    default_slide_count: int = 6


class TextModelConfig(BaseModel):
    provider: Literal["openai", "gemini"] = "openai"
    base_url: str = "https://yunwu.ai/v1"
    api_key: str = ""
    model: str = "gpt-5.2-chat-latest"


class EditableModelConfig(BaseModel):
    provider: Literal["openai", "gemini"] = "openai"
    base_url: str = "https://yunwu.ai/v1"
    api_key: str = ""
    model: str = "gemini-3.1-pro-preview"
    prompt_file: str = ""
    browser_path: str = ""
    download_timeout_ms: int = 180000
    max_tokens: int = 1000000
    max_attempts: int = 3
    sleep_seconds: float = 1.0
    asset_backend: str = "edit"
    disable_asset_reuse: bool = False


class ImageModelConfig(BaseModel):
    provider: Literal["openai", "gemini", "http"] = "http"
    base_url: str = "https://grsai.dakka.com.cn/v1/draw/completions"
    api_key: str = ""
    model: str = "gpt-image-2"
    size: str = "4K"
    variants: int = 1
    timeout: int = 300
    retries: int = 2
    max_workers: int = 20


class ModelConfigGroup(BaseModel):
    text: TextModelConfig = Field(default_factory=TextModelConfig)
    editable: EditableModelConfig = Field(default_factory=EditableModelConfig)
    image: ImageModelConfig = Field(default_factory=ImageModelConfig)


class MineruConfig(BaseModel):
    base_url: str = "https://mineru.net/api/v4"
    api_key: str = ""
    model_version: str = "vlm"
    language: str = "ch"
    enable_formula: bool = True
    enable_table: bool = True
    is_ocr: bool = True
    poll_interval_seconds: float = 2.0
    timeout_seconds: int = 300
    max_refine_depth: int = 2


class Settings(BaseModel):
    app: AppConfig = Field(default_factory=AppConfig)
    models: ModelConfigGroup = Field(default_factory=ModelConfigGroup)
    mineru: MineruConfig = Field(default_factory=MineruConfig)

    @property
    def output_root(self) -> str:
        return self.app.output_root

    @property
    def default_slide_count(self) -> int:
        return self.app.default_slide_count

    @property
    def text_provider(self) -> str:
        return self.models.text.provider

    @property
    def text_base_url(self) -> str:
        return self.models.text.base_url

    @property
    def text_api_key(self) -> str:
        return self.models.text.api_key

    @property
    def text_model(self) -> str:
        return self.models.text.model

    @property
    def editable_ppt_provider(self) -> str:
        return self.models.editable.provider

    @property
    def editable_ppt_base_url(self) -> str:
        return self.models.editable.base_url

    @property
    def editable_ppt_api_key(self) -> str:
        return self.models.editable.api_key

    @property
    def editable_ppt_model(self) -> str:
        return self.models.editable.model

    @property
    def editable_ppt_prompt_file(self) -> str:
        return self.models.editable.prompt_file

    @property
    def editable_ppt_browser_path(self) -> str:
        return self.models.editable.browser_path

    @property
    def editable_ppt_download_timeout_ms(self) -> int:
        return self.models.editable.download_timeout_ms

    @property
    def editable_ppt_max_tokens(self) -> int:
        return self.models.editable.max_tokens

    @property
    def editable_ppt_max_attempts(self) -> int:
        return self.models.editable.max_attempts

    @property
    def editable_ppt_sleep_seconds(self) -> float:
        return self.models.editable.sleep_seconds

    @property
    def editable_ppt_asset_backend(self) -> str:
        return self.models.editable.asset_backend

    @property
    def editable_ppt_disable_asset_reuse(self) -> bool:
        return self.models.editable.disable_asset_reuse

    @property
    def image_provider(self) -> str:
        return self.models.image.provider

    @property
    def image_api_url(self) -> str:
        return self.models.image.base_url

    @property
    def image_api_key(self) -> str:
        return self.models.image.api_key

    @property
    def image_model(self) -> str:
        return self.models.image.model

    @property
    def image_size(self) -> str:
        return self.models.image.size

    @property
    def image_variants(self) -> int:
        return self.models.image.variants

    @property
    def image_timeout(self) -> int:
        return self.models.image.timeout

    @property
    def image_retries(self) -> int:
        return self.models.image.retries

    @property
    def image_max_workers(self) -> int:
        return self.models.image.max_workers

    @property
    def openai_base_url(self) -> str:
        return self.text_base_url

    @property
    def resolved_image_key(self) -> str:
        return self.models.image.api_key or self.models.text.api_key

    @property
    def resolved_editable_base_url(self) -> str:
        return self.models.editable.base_url or self.models.text.base_url

    @property
    def resolved_editable_api_key(self) -> str:
        return self.models.editable.api_key or self.models.text.api_key

    @property
    def mineru_base_url(self) -> str:
        return self.mineru.base_url

    @property
    def mineru_api_key(self) -> str:
        return self.mineru.api_key

    @property
    def mineru_model_version(self) -> str:
        return self.mineru.model_version

    @property
    def mineru_language(self) -> str:
        return self.mineru.language

    @property
    def mineru_enable_formula(self) -> bool:
        return self.mineru.enable_formula

    @property
    def mineru_enable_table(self) -> bool:
        return self.mineru.enable_table

    @property
    def mineru_is_ocr(self) -> bool:
        return self.mineru.is_ocr

    @property
    def mineru_poll_interval_seconds(self) -> float:
        return self.mineru.poll_interval_seconds

    @property
    def mineru_timeout_seconds(self) -> int:
        return self.mineru.timeout_seconds

    @property
    def mineru_max_refine_depth(self) -> int:
        return self.mineru.max_refine_depth

    @property
    def resolved_mineru_base_url(self) -> str:
        return self.mineru.base_url or "https://mineru.net/api/v4"

    @property
    def resolved_mineru_api_key(self) -> str:
        return self.mineru.api_key or self.models.editable.api_key or self.models.text.api_key


def _read_yaml_config(path: Path) -> dict[str, Any]:
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if not isinstance(raw, dict):
        raise ValueError(f"Config file must be a YAML object: {path}")
    return raw

def _load_raw_config(config_file: Optional[str]) -> dict[str, Any]:
    if config_file:
        config_path = Path(config_file).expanduser().resolve()
        if not config_path.exists():
            raise FileNotFoundError(f"Config file not found: {config_path}")
        if config_path.suffix.lower() not in {".yaml", ".yml"}:
            raise ValueError("Config file must be a .yaml or .yml file.")
        return _read_yaml_config(config_path)

    if DEFAULT_CONFIG_FILE.exists():
        return _read_yaml_config(DEFAULT_CONFIG_FILE)

    return {}


@lru_cache(maxsize=1)
def get_settings() -> Settings:
    return load_settings()


def load_settings(config_file: Optional[str] = None) -> Settings:
    return Settings.model_validate(_load_raw_config(config_file))
