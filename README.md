# EditDeck

<p align="center">An end-to-end pipeline from requirement text to slide images, standard PPTX, and editable PPTX.</p>

<p align="center">
  <img src="https://img.shields.io/badge/Python-FastAPI-blue?style=flat-square" alt="Python FastAPI" />
  <img src="https://img.shields.io/badge/Config-YAML-0EA5E9?style=flat-square" alt="YAML Config" />
  <img src="https://img.shields.io/badge/Workflow-Web%20%2B%20CLI-10B981?style=flat-square" alt="Web and CLI" />
  <img src="https://img.shields.io/badge/License-MIT-black?style=flat-square" alt="MIT License" />
</p>

<p align="center">
  <a href="#why">Why</a> ·
  <a href="#highlights">Highlights</a> ·
  <a href="#case-study">Case Study</a> ·
  <a href="#quick-start">Quick Start</a> ·
  <a href="#usage">Usage</a> ·
  <a href="#configuration">Configuration</a> ·
  <a href="#faq">FAQ</a> ·
  <a href="./README_CN.md">中文文档</a>
</p>

---

<a id="why"></a>

## Why

Making slides usually means bouncing between too many disconnected steps: shaping the story, drafting page content, generating visuals, exporting a deck, and then rebuilding everything again if you need a truly editable presentation.

EditDeck turns that fragmented process into one continuous workflow:

- Start from a plain-language requirement and turn it into a structured PPT outline with page-level content
- Render every slide into polished visual images and package them into a standard `pptx`
- Continue from the generated run directory or from existing slide images to rebuild an editable `pptx`
- Coordinate text models, image models, editable deck generation, and MinerU parsing through one unified `YAML` config
- Expose the same workflow through Web UI, CLI, and HTTP API so it fits both hands-on usage and system integration

If your ideal workflow is "get the visual draft fast, then keep going until the deck becomes editable and usable," this project is built for exactly that.

<a id="highlights"></a>

## Highlights

- Single config entry point: the project reads only [config/app.yaml](./config/app.yaml) by default
- Dual workflow support: generate from scratch or re-generate an editable PPT from existing images
- Complete editable pipeline: image parsing, element extraction, placeholder matching, and browser export are all wired up
- Better cross-platform support: the browser path can be left empty — at runtime it auto-detects from explicit arguments, environment variables, and the system `PATH`
- Straightforward overrides: both CLI arguments and Web/API request parameters can override the config file at runtime

## Workflow

```text
Requirement
  -> Outline / Page Content
  -> Slide Images
  -> Standard PPTX
  -> MinerU Asset Parsing
  -> Browser-side Placeholder Matching
  -> Editable PPTX
```

<a id="case-study"></a>

## Case Study

Following the side-by-side comparison style used by PPTAgent, each case is shown in two columns: the left side is the image-only result, and the right side is the fully editable reconstruction.

### Case 1

**Prompt**

> 做一份 打造高质量PPT汇报的系统方法 6页

<table>
  <tr>
    <th width="50%">Image-only Deck</th>
    <th width="50%">Fully Editable Deck</th>
  </tr>
  <tr>
    <td><img src="./case-study/contact-sheets/demo1.png" alt="Case 1 image-only deck" /></td>
    <td><img src="./case-study/contact-sheets/demo1_buke.png" alt="Case 1 fully editable deck" /></td>
  </tr>
</table>

### Case 2

**Prompt**

> 做一份《AI 客服知识库升级方案》PPT，面向企业数字化与客服平台主管，围绕现状痛点、升级目标、知识中台架构、问答流程优化、运营指标与实施计划展开，适合六页呈现。

**Style**

> PPT 整体呈现 16:9 宽屏深色科技风，主色以深海军蓝、冷青蓝、电光蓝为核心，辅以少量高亮青色作为数据强调色，背景以深蓝黑渐变、细密网格、弱发光线条和玻璃质感面板构成。字体采用思源黑体 / 苹方 / 微软雅黑体系，标题更厚重，正文更克制，层级鲜明。页面强调中轴对齐与模块化网格系统，常用大标题横向锚点、分栏数据卡、发光描边图表、半透明信息面板与线性科技图标。整体气质冷静、专业、偏未来感，但必须保持高可读性，避免赛博朋克式杂乱，避免大面积紫色，强调企业级 AI 产品汇报的理性秩序。

<table>
  <tr>
    <th width="50%">Image-only Deck</th>
    <th width="50%">Fully Editable Deck</th>
  </tr>
  <tr>
    <td><img src="./case-study/contact-sheets/demo2.png" alt="Case 2 image-only deck" /></td>
    <td><img src="./case-study/contact-sheets/demo2_buke.png" alt="Case 2 fully editable deck" /></td>
  </tr>
</table>

### Case 3

**Prompt**

> 做一份《企业 Copilot 落地的双引擎实施路线图》PPT，面向集团管理层汇报，围绕数据治理引擎与业务应用引擎两条主线，讲清建设背景、核心痛点、总体架构、分阶段路线图、试点场景、投入产出与风险控制，适合六页呈现。

**Style**

> PPT 整体呈现16:9 宽屏学术商务风，以微软红蓝为主色调，搭配浅灰与蓝灰底色，采用思源黑体 / 苹方 / 微软雅黑体系，层级清晰、信息密度适中。页面采用顶部标题栏 + 左侧纵向时间轴 + 右侧斜向分区双引擎的差异化构图，所有卡片统一 18–22px 圆角与轻量阴影，模块标题使用胶囊色块，重要容器配以红蓝虚线外框，搭配统一线宽线性图标强化视觉规范。

<table>
  <tr>
    <th width="50%">Image-only Deck</th>
    <th width="50%">Fully Editable Deck</th>
  </tr>
  <tr>
    <td><img src="./case-study/contact-sheets/demo3.png" alt="Case 3 image-only deck" /></td>
    <td><img src="./case-study/contact-sheets/demo3_buke.png" alt="Case 3 fully editable deck" /></td>
  </tr>
</table>

<a id="quick-start"></a>

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Edit Configuration

Edit [config/app.yaml](./config/app.yaml).

- The `api_key` fields are left empty by default in the project template
- You can keep using the `base_url` addresses already in the file
- See [config/README.md](./config/README.md) for a more complete field reference

### 3. Choose How to Run

Start the web server:

```bash
uvicorn webapp.main:app --host 0.0.0.0 --port 8000 --reload
```

Open in your browser:

```text
http://127.0.0.1:8000/
```

Or use the CLI directly:

```bash
python -m app.cli generate "Create an AI office productivity improvement plan"
```

<a id="usage"></a>

## Usage

### Web

The web entry point is provided by [webapp/main.py](./webapp/main.py), ideal for filling in requirements, styles, and runtime parameters directly on the page.

### CLI

Generate images and a standard PPT only:

```bash
python -m app.cli generate "Create an AI office productivity improvement plan" \
  --slide-count auto \
  --export-mode both
```

Generate a standard PPT and then continue to produce an editable PPT:

```bash
python -m app.cli generate "Create an AI office productivity improvement plan" \
  --editable-ppt \
  -edit
```

Continue generating an editable PPT from an existing run directory:

```bash
python -m app.cli editable \
  --run-dir ./generated/<run_id> \
  --output-dir ./generated/<run_id>/editable_deck \
  -edit
```

Generate an editable PPT directly from existing images:

```bash
python -m app.cli editable \
  --image ./generated/run_xxx/slide_01.png \
  --image ./generated/run_xxx/slide_02.png \
  --output-dir ./generated/run_xxx/editable_deck \
  -edit
```

Common parameters:

- `--config-file`: specify a config file; defaults to `config/app.yaml`
- `--style-description`: specify style via text
- `--style-template`: specify style via image
- `--editable-ppt`: continue generating an editable PPT after image generation
- `-edit` / `--edit`: enable the currently available editable asset matching backend
- `--mineru-api-key`: override `mineru.api_key` as needed
- `--force-reextract-assets`: force re-extraction of elements
- `--disable-asset-reuse`: prevent a single asset from being reused across multiple placeholders

Notes:

- `--style-description` and `--style-template` are mutually exclusive
- CLI arguments take priority over `YAML` configuration

## HTTP API

Main endpoints:

- `GET /api/health`: health check
- `POST /api/generate`: synchronous generation
- `POST /api/generate/start`: asynchronous generation
- `POST /api/editable/start`: start an editable PPT task from an existing `run_id`
- `GET /api/generate/status/{job_id}`: query async task status

To produce an editable PPT directly during the generation phase, include the following in your request:

- `generate_editable_ppt=true`
- `asset_backend=edit`

When `config/app.yaml` does not have a usable `mineru.api_key`, you need to pass `mineru_api_key` explicitly in the request.

<a id="configuration"></a>

## Configuration

The project keeps a single main config file:

```text
config/app.yaml
```

Config sections:

- `app`: output directory and default slide count
- `models.text`: models for outline and copy generation
- `models.editable`: the editable PPT generation pipeline
- `models.image`: image generation model
- `mineru`: page element parsing and asset extraction

Useful fallback rules:

- If `models.image.api_key` is empty, it falls back to `models.text.api_key`
- If `models.editable.base_url` is empty, it falls back to `models.text.base_url`
- If `models.editable.api_key` is empty, it falls back to `models.text.api_key`
- If `mineru.api_key` is empty, it further falls back to `models.editable.api_key` and then `models.text.api_key`
- `models.editable.browser_path` can be left empty — at runtime it tries explicit arguments, environment variables, and the system `PATH`

For a complete example and field reference, see [config/README.md](./config/README.md).

## Output

Each run writes results under `generated/<run_id>/`. A typical directory structure:

```text
generated/<run_id>/
├─ slide_01.png
├─ slide_02.png
├─ ...
├─ *.pptx
├─ editable_deck/
│  ├─ editable_deck.pptx
│  ├─ result.json
│  └─ ...
└─ logs/
```

The editable pipeline also leaves behind these intermediate artifacts for debugging:

- `edit_assets/`
- `attempt_01/`
- `filled_preview/`
- `browser_asset_manifest.json`

## Project Structure

```text
.
├─ app/
│  ├─ cli.py
│  ├─ pipeline.py
│  ├─ settings.py
│  └─ editable_ppt/
├─ webapp/
│  ├─ main.py
│  └─ static/
├─ config/
├─ scripts/
├─ generated/
└─ requirements.txt
```

Core files:

- [app/cli.py](./app/cli.py)
- [app/pipeline.py](./app/pipeline.py)
- [app/settings.py](./app/settings.py)
- [app/editable_ppt/service.py](./app/editable_ppt/service.py)
- [app/editable_ppt/mineru_assets.py](./app/editable_ppt/mineru_assets.py)
- [webapp/main.py](./webapp/main.py)

<a id="faq"></a>

## FAQ

### Editable PPT reports a missing key

Check the following in order:

- `mineru.api_key` in [config/app.yaml](./config/app.yaml)
- `--mineru-api-key` in CLI arguments
- `mineru_api_key` in Web / API requests

### Browser execution or download fails

Troubleshoot in this order:

- First, leave `models.editable.browser_path` empty
- If you need to specify a browser explicitly, pass `--editable-browser-path`
- Or set one of these environment variables: `EDITABLE_PPT_BROWSER_PATH`, `CHROME_PATH`, `GOOGLE_CHROME_BIN`, `CHROMIUM_PATH`, `BROWSER_PATH`
- If no browser is available on the system, run `playwright install chromium`

### Placeholder replacement looks off

Try:

- Increase `mineru.max_refine_depth`
- Enable `--force-reextract-assets`
- Enable `--disable-asset-reuse`

### Just want to reuse existing assets

Use `--assets-json`, though this is currently best suited for directly specifying an existing `assets.json` in single-image mode.
