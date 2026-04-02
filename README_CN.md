# EditDeck

<p align="center">从需求文本到 PPT 图片、普通 PPT、可编辑 PPT 的一体化生成流水线。</p>

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
  <a href="#faq">FAQ</a>
</p>

---

<a id="why"></a>

## Why

做 PPT 往往不是难在某一步，而是难在流程太碎: 想清楚结构、补页面内容、生成视觉稿、导出文档、最后还要为了“可编辑”再重新搭一遍。

EditDeck 把这条原本零散、反复切换的链路，收拢成一套可以连续推进的工作流：

- 从自然语言需求出发，自动生成结构化的大纲和逐页内容
- 批量渲染每一页视觉稿，并打包成普通 `pptx`
- 可以接着基于生成结果目录，或直接基于已有页面图片，继续重建可编辑 `pptx`
- 用统一的 `YAML` 配置同时管理文本模型、图片模型、可编辑链路和 MinerU 解析能力
- 同一套能力同时提供 Web、CLI 和 HTTP API，既适合直接使用，也方便接入更大的系统

如果你想要的是“先尽快把视觉稿跑出来，再一路推进到真正可编辑、可交付的演示文档”，这套流程就是为这个目标设计的。

<a id="highlights"></a>

## Highlights

- 单配置入口：项目默认只读取 [config/app.yaml](/E:/xinda_agent2206/config/app.yaml)
- 双工作流支持：既可以从需求直接生成，也可以对已有图片二次生成可编辑 PPT
- 可编辑链路完整：图片解析、元素抽取、占位匹配、浏览器导出都已经串起来
- 跨平台更友好：浏览器路径可以留空，运行时会按传入路径、环境变量、系统 `PATH` 自动探测
- 覆盖方式直接：CLI 参数和 Web/API 请求参数都可以在运行时覆盖配置文件

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

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 修改配置

编辑 [config/app.yaml](/E:/xinda_agent2206/config/app.yaml)。

- 项目模板中的 `api_key` 默认留空
- `base_url` 可以继续使用当前文件里的地址
- 更完整的字段说明见 [config/README.md](/E:/xinda_agent2206/config/README.md)

### 3. 选择运行方式

启动 Web 服务：

```bash
uvicorn webapp.main:app --host 0.0.0.0 --port 8000 --reload
```

浏览器访问：

```text
http://127.0.0.1:8000/
```

或者直接使用 CLI：

```bash
python -m app.cli generate "做一份 AI 办公效率提升方案"
```

<a id="usage"></a>

## Usage

### Web

Web 入口由 [webapp/main.py](/E:/xinda_agent2206/webapp/main.py) 提供，适合直接在页面里填写需求、风格和运行参数。

### CLI

只生成图片和普通 PPT：

```bash
python -m app.cli generate "做一份 AI 办公效率提升方案" \
  --slide-count auto \
  --export-mode both
```

生成普通 PPT 后继续输出可编辑 PPT：

```bash
python -m app.cli generate "做一份 AI 办公效率提升方案" \
  --editable-ppt \
  -edit
```

基于已有运行目录继续生成可编辑 PPT：

```bash
python -m app.cli editable \
  --run-dir ./generated/<run_id> \
  --output-dir ./generated/<run_id>/editable_deck \
  -edit
```

基于现有图片直接生成可编辑 PPT：

```bash
python -m app.cli editable \
  --image ./generated/run_xxx/slide_01.png \
  --image ./generated/run_xxx/slide_02.png \
  --output-dir ./generated/run_xxx/editable_deck \
  -edit
```

常用参数：

- `--config-file`：指定配置文件，默认读取 `config/app.yaml`
- `--style-description`：用文字指定风格
- `--style-template`：用图片指定风格
- `--editable-ppt`：在生成图片后继续生成可编辑 PPT
- `-edit` / `--edit`：启用当前可用的可编辑资产匹配后端
- `--mineru-api-key`：按需覆盖 `mineru.api_key`
- `--force-reextract-assets`：强制重新抽取元素
- `--disable-asset-reuse`：禁止一个素材复用到多个 `PH`

说明：

- `--style-description` 和 `--style-template` 互斥
- CLI 传参优先级高于 `YAML` 配置

## HTTP API

主要接口：

- `GET /api/health`：健康检查
- `POST /api/generate`：同步生成
- `POST /api/generate/start`：异步生成
- `POST /api/editable/start`：基于已有 `run_id` 启动可编辑 PPT 任务
- `GET /api/generate/status/{job_id}`：查询异步任务状态

如果希望在生成阶段直接产出可编辑 PPT，请在请求里额外传入：

- `generate_editable_ppt=true`
- `asset_backend=edit`

当 `config/app.yaml` 里没有可用的 `mineru.api_key` 时，需要通过请求显式传入 `mineru_api_key`。

<a id="configuration"></a>

## Configuration

项目只保留一个主配置文件：

```text
config/app.yaml
```

配置块说明：

- `app`：输出目录和默认页数
- `models.text`：大纲、文案等文本生成模型
- `models.editable`：可编辑 PPT 生成链路
- `models.image`：图片生成模型
- `mineru`：页面元素解析与资产抽取

几个实用规则：

- `models.image.api_key` 为空时，会回退使用 `models.text.api_key`
- `models.editable.base_url` 为空时，会回退使用 `models.text.base_url`
- `models.editable.api_key` 为空时，会回退使用 `models.text.api_key`
- `mineru.api_key` 为空时，会继续回退尝试 `models.editable.api_key` 和 `models.text.api_key`
- `models.editable.browser_path` 可以留空，运行时会自动尝试显式传参、环境变量和系统 `PATH`

完整示例和字段参考见 [config/README.md](/E:/xinda_agent2206/config/README.md)。

## Output

每次运行会在 `generated/<run_id>/` 下写出结果。常见目录结构如下：

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

可编辑链路通常还会留下这些中间结果，方便排查问题：

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

核心文件：

- [app/cli.py](/E:/xinda_agent2206/app/cli.py)
- [app/pipeline.py](/E:/xinda_agent2206/app/pipeline.py)
- [app/settings.py](/E:/xinda_agent2206/app/settings.py)
- [app/editable_ppt/service.py](/E:/xinda_agent2206/app/editable_ppt/service.py)
- [app/editable_ppt/mineru_assets.py](/E:/xinda_agent2206/app/editable_ppt/mineru_assets.py)
- [webapp/main.py](/E:/xinda_agent2206/webapp/main.py)

<a id="faq"></a>

## FAQ

### 可编辑 PPT 提示缺少 key

优先检查：

- [config/app.yaml](/E:/xinda_agent2206/config/app.yaml) 里的 `mineru.api_key`
- CLI 参数里的 `--mineru-api-key`
- Web / API 请求里的 `mineru_api_key`

### 浏览器执行失败或下载失败

可以按这个顺序排查：

- 先让 `models.editable.browser_path` 保持空值
- 如果需要显式指定浏览器，再传 `--editable-browser-path`
- 或者配置 `EDITABLE_PPT_BROWSER_PATH`、`CHROME_PATH`、`GOOGLE_CHROME_BIN`、`CHROMIUM_PATH`、`BROWSER_PATH`
- 如果系统里没有可用浏览器，执行 `playwright install chromium`

### 占位替换效果不理想

可以尝试：

- 提高 `mineru.max_refine_depth`
- 开启 `--force-reextract-assets`
- 开启 `--disable-asset-reuse`

### 只想复用已有素材

可以使用 `--assets-json`，但当前更适合单图片模式下直接指定已有 `assets.json`。
