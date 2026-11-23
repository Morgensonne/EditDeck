# Configuration Reference

项目默认只读取一个配置文件：

```text
config/app.yaml
```

如果没有额外传入 CLI 参数或 Web/API 运行参数，系统会直接按这个文件执行。

## Override Order

运行时配置优先级如下：

```text
CLI / Web / API overrides
  > config/app.yaml
  > code defaults
```

这意味着：

- 适合长期保存的内容，写进 `config/app.yaml`
- 只想临时切换模型或 key，可以在 CLI 或 Web/API 请求里覆盖
- 模板仓库里的 `api_key` 可以保持空字符串

## Full Example

```yaml
app:
  output_root: generated
  default_slide_count: 6

models:
  text:
    provider: openai
    base_url: https://yunwu.ai/v1
    api_key: ""
    model: gpt-5.2-chat-latest

  editable:
    provider: openai
    base_url: https://grsaiapi.com/v1
    api_key: ""
    model: gemini-3.1-pro
    prompt_file: ""
    browser_path: ""
    download_timeout_ms: 180000
    max_tokens: 1000000
    max_attempts: 3
    sleep_seconds: 1.0
    asset_backend: edit
    disable_asset_reuse: false

  image:
    provider: http
    base_url: https://grsai.dakka.com.cn/v1/draw/nano-banana
    api_key: ""
    model: nano-banana-pro
    size: 4K
    variants: 1
    timeout: 900
    retries: 2
    max_workers: 20

mineru:
  base_url: https://mineru.net/api/v4
  api_key: ""
  model_version: vlm
  language: ch
  enable_formula: true
  enable_table: true
  is_ocr: true
  poll_interval_seconds: 2.0
  timeout_seconds: 300
  max_refine_depth: 2
```

## Sections

### `app`

| Field | Description | Default |
| --- | --- | --- |
| `output_root` | 结果输出目录 | `generated` |
| `default_slide_count` | 默认页数 | `6` |

### `models.text`

用于生成大纲、页面文案和主流程文本内容。

| Field | Description | Notes |
| --- | --- | --- |
| `provider` | 文本模型提供方 | 支持 `openai`、`gemini` |
| `base_url` | 文本接口地址 | 也会作为部分回退来源 |
| `api_key` | 文本模型 key | 可以先留空 |
| `model` | 文本模型名 | 由调用方自行替换 |

### `models.editable`

用于可编辑 PPT 链路，包括代码生成、浏览器执行和导出。

| Field | Description | Notes |
| --- | --- | --- |
| `provider` | 可编辑模型提供方 | 支持 `openai`、`gemini` |
| `base_url` | 可编辑模型接口地址 | 为空时回退到 `models.text.base_url` |
| `api_key` | 可编辑模型 key | 为空时回退到 `models.text.api_key` |
| `model` | 可编辑模型名 | 例如 Gemini / OpenAI 兼容模型 |
| `prompt_file` | 自定义 prompt 文件路径 | 为空则使用默认 prompt |
| `browser_path` | 浏览器可执行文件路径 | 可以留空，运行时自动探测 |
| `download_timeout_ms` | 浏览器下载超时 | 单位毫秒 |
| `max_tokens` | 可编辑模型最大输出 token | 影响生成代码长度 |
| `max_attempts` | 单页最大重试次数 | 用于失败恢复 |
| `sleep_seconds` | 重试间隔 | 单位秒 |
| `asset_backend` | 资产匹配后端 | 当前固定使用 `edit` |
| `disable_asset_reuse` | 是否禁止素材复用 | `true` 时一个素材不会匹配多个 `PH` |

浏览器解析顺序：

- 显式传入路径
- 环境变量：`EDITABLE_PPT_BROWSER_PATH`、`CHROME_PATH`、`GOOGLE_CHROME_BIN`、`CHROMIUM_PATH`、`BROWSER_PATH`
- 系统 `PATH` 中的 Chrome / Chromium / Edge / Brave
- 如果没有现成浏览器，可安装 Playwright Chromium

### `models.image`

用于页面图片生成。

| Field | Description | Notes |
| --- | --- | --- |
| `provider` | Image provider | `http`, `openai`, `gemini`, or `minimax` |
| `base_url` | 图片接口地址 | 保留当前服务地址即可 |
| `api_key` | Image API key | Set a dedicated image key for H3. Other image providers can fall back to `models.text.api_key`. |
| `model` | 生图模型名 | 例如 `nano-banana-pro` |
| `size` | 图片尺寸 | 由后端解释 |
| `variants` | 单次生成张数 | 通常保持 `1` |
| `timeout` | 单次请求超时 | 单位秒 |
| `retries` | 重试次数 | 失败时使用 |
| `max_workers` | 并发数 | 控制批量生成速度 |

#### MiniMax H3 slide images

Set the following fields under `models.image` in `config/app.yaml` and supply the image key through `--image-api-key` or `models.image.api_key`.

```yaml
provider: minimax
base_url: https://api.minimax.io/v1/images/generations
model: MiniMax-H3-Image
size: 4K
variants: 1
```

For a China-region key, use `https://api.minimax.cn/v1/images/generations`. Both regions also support `MiniMax-H3-Image-Highspeed`. The area presets `1K`, `2K` and `4K` request a widescreen slide. An explicit `WIDTHxHEIGHT` requests that size directly. Each slide uses the first successful image returned. Keep `variants: 1` to generate one image per slide.

H3 generation waits at least 600 seconds and sends each generation request once because a timed-out request can still complete and incur a charge.

### `mineru`

用于页面元素解析和资产抽取，是可编辑 PPT 能否补全的关键配置。

| Field | Description | Notes |
| --- | --- | --- |
| `base_url` | MinerU 接口地址 | 默认 `https://mineru.net/api/v4` |
| `api_key` | MinerU key | 为空时回退到 `models.editable.api_key`，再回退到 `models.text.api_key` |
| `model_version` | MinerU 模型版本 | 默认 `vlm` |
| `language` | 解析语言 | 默认 `ch` |
| `enable_formula` | 是否启用公式提取 | 布尔值 |
| `enable_table` | 是否启用表格提取 | 布尔值 |
| `is_ocr` | 是否启用 OCR | 布尔值 |
| `poll_interval_seconds` | 轮询间隔 | 单位秒 |
| `timeout_seconds` | 任务超时 | 单位秒 |
| `max_refine_depth` | 最大细化深度 | 值越大通常越慢 |

## Practical Tips

- 仓库模板里的 `api_key` 建议继续留空，部署时再填自己的值。
- 如果你希望配置更便于跨平台迁移，不要写死 Windows 专属浏览器路径，`browser_path` 留空通常更稳。
- 临时实验某个模型时，优先用 CLI 参数或 Web/API 覆盖，不必改动主配置文件。
- 需要生成可编辑 PPT 时，`mineru` 相关配置比普通生图链路更关键。
