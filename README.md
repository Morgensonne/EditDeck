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
  <a href="#quick-start">Quick Start</a> ·
  <a href="#usage">Usage</a> ·
  <a href="#configuration">Configuration</a> ·
  <a href="#faq">FAQ</a> ·
  <a href="./README_CN.md">中文文档</a>
</p>

---

<a id="why"></a>

## Why

This project chains together several otherwise scattered tasks into one complete pipeline:

- Automatically generate a PPT outline and page content from requirements
- Batch-render each slide as an image and export a standard `pptx`
- Continue from an existing run directory or existing images to produce an editable `pptx`
- Manage text models, editable models, image models, and MinerU through a unified `YAML` config
- Support Web UI, CLI, and HTTP API simultaneously

If what you want is "get a visual draft fast, then turn it into an editable presentation," this workflow should feel natural.

<a id="highlights"></a>

## Highlights

- Single config entry point: the project reads only [config/app.yaml](/E:/xinda_agent2206/config/app.yaml) by default
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

<a id="quick-start"></a>

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Edit Configuration

Edit [config/app.yaml](/E:/xinda_agent2206/config/app.yaml).

- The `api_key` fields are left empty by default in the project template
- You can keep using the `base_url` addresses already in the file
- See [config/README.md](/E:/xinda_agent2206/config/README.md) for a more complete field reference

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

The web entry point is provided by [webapp/main.py](/E:/xinda_agent2206/webapp/main.py), ideal for filling in requirements, styles, and runtime parameters directly on the page.

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

For a complete example and field reference, see [config/README.md](/E:/xinda_agent2206/config/README.md).

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

- [app/cli.py](/E:/xinda_agent2206/app/cli.py)
- [app/pipeline.py](/E:/xinda_agent2206/app/pipeline.py)
- [app/settings.py](/E:/xinda_agent2206/app/settings.py)
- [app/editable_ppt/service.py](/E:/xinda_agent2206/app/editable_ppt/service.py)
- [app/editable_ppt/mineru_assets.py](/E:/xinda_agent2206/app/editable_ppt/mineru_assets.py)
- [webapp/main.py](/E:/xinda_agent2206/webapp/main.py)

<a id="faq"></a>

## FAQ

### Editable PPT reports a missing key

Check the following in order:

- `mineru.api_key` in [config/app.yaml](/E:/xinda_agent2206/config/app.yaml)
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