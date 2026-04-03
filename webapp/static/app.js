const form = document.getElementById("workflow-form");
const statusBox = document.getElementById("status");
const defaultsText = document.getElementById("defaults-text");
const progressCard = document.getElementById("progress-card");
const progressStep = document.getElementById("progress-step");
const progressPercent = document.getElementById("progress-percent");
const progressFill = document.getElementById("progress-fill");
const progressDetail = document.getElementById("progress-detail");
const progressSlide = document.getElementById("progress-slide");
const styleTemplateInput = form.querySelector('input[name="style_template"]');
const styleTemplateBase64Input = form.querySelector('input[name="style_template_base64"]');

const generateFields = document.getElementById("generate-fields");
const replicaFields = document.getElementById("replica-fields");
const prepareBtn = document.getElementById("prepare-btn");
const replicaBtn = document.getElementById("replica-btn");
const saveEditsBtn = document.getElementById("save-edits-btn");
const renderBtn = document.getElementById("render-btn");
const rerenderBtn = document.getElementById("rerender-btn");
const editableBtn = document.getElementById("editable-btn");
const editableSelectedBtn = document.getElementById("editable-selected-btn");

const editorCard = document.getElementById("editor-card");
const resultCard = document.getElementById("result-card");
const editablePreviewCard = document.getElementById("editable-preview-card");
const sessionIdInput = document.getElementById("session-id-input");
const deckTitleInput = document.getElementById("deck-title-input");
const stylePromptInput = document.getElementById("style-prompt-input");
const outlineEditor = document.getElementById("outline-editor");
const slidesGrid = document.getElementById("slides-grid");
const editableGrid = document.getElementById("editable-grid");
const pptxLink = document.getElementById("pptx-link");
const editableLink = document.getElementById("editable-link");

const STEP_LABELS = {
  queued: "排队中",
  prepare: "准备中",
  slide_count: "页数规划",
  style: "风格生成",
  outline: "大纲生成",
  prompt_generation: "Prompt生成",
  image_generation: "图片生成",
  packaging: "PPT打包",
  editable_prepare: "可编辑准备",
  editable_assets: "素材匹配",
  editable_codegen: "代码生成",
  editable_render: "可编辑渲染",
  editable_packaging: "可编辑打包",
  completed: "已完成",
  failed: "失败",
};

const MODEL_FIELDS = [
  "base_url",
  "image_api_url",
  "text_api_key",
  "image_api_key",
  "text_model",
  "image_model",
  "mineru_base_url",
  "mineru_api_key",
  "mineru_model_version",
  "mineru_language",
  "mineru_enable_formula",
  "mineru_enable_table",
  "mineru_is_ocr",
  "mineru_poll_interval_seconds",
  "mineru_timeout_seconds",
  "mineru_max_refine_depth",
  "force_reextract_assets",
  "disable_asset_reuse",
];

const EDITABLE_FIELDS = [
  "editable_base_url",
  "editable_api_key",
  "editable_model",
  "editable_prompt_file",
  "editable_browser_path",
  "editable_download_timeout_ms",
  "editable_max_tokens",
  "editable_max_attempts",
  "editable_sleep_seconds",
  "assets_dir",
  "asset_backend",
  "mineru_base_url",
  "mineru_api_key",
  "mineru_model_version",
  "mineru_language",
  "mineru_enable_formula",
  "mineru_enable_table",
  "mineru_is_ocr",
  "mineru_poll_interval_seconds",
  "mineru_timeout_seconds",
  "mineru_max_refine_depth",
  "force_reextract_assets",
  "disable_asset_reuse",
];

const state = {
  mode: "generate",
  sessionId: "",
  pollToken: 0,
  selectedPages: new Set(),
  outlineFingerprint: "",
  slidesFingerprint: "",
};

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.classList.remove("hidden");
  statusBox.classList.toggle("error", isError);
}

function resetProgress() {
  progressCard.classList.remove("hidden");
  progressStep.textContent = "准备中...";
  progressPercent.textContent = "0%";
  progressFill.style.width = "0%";
  progressDetail.textContent = "";
  progressSlide.classList.add("hidden");
  progressSlide.textContent = "";
}

function updateProgress(job) {
  progressCard.classList.remove("hidden");
  const progress = Number(job.progress || 0);
  progressStep.textContent = STEP_LABELS[job.step] || job.step || "处理中";
  progressPercent.textContent = `${progress}%`;
  progressFill.style.width = `${Math.min(100, Math.max(0, progress))}%`;
  progressDetail.textContent = job.message || "";
  const current = Number(job.current_slide || 0);
  const total = Number(job.total_slides || 0);
  if (total > 0) {
    progressSlide.classList.remove("hidden");
    progressSlide.textContent = `进度：${current}/${total}`;
  } else {
    progressSlide.classList.add("hidden");
    progressSlide.textContent = "";
  }
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve((reader.result || "").toString());
    reader.onerror = () => reject(new Error("模板图读取失败，请重新选择文件。"));
    reader.readAsDataURL(file);
  });
}

function getFieldNodes(name) {
  return form.querySelectorAll(`[name="${name}"]`);
}

function appendOptionalField(formData, name) {
  const nodes = getFieldNodes(name);
  if (!nodes.length) return;
  const first = nodes[0];
  if (first.type === "checkbox") {
    const checked = Array.from(nodes).some((node) => node.checked);
    if (checked) formData.set(name, "true");
    return;
  }
  if (first.type === "radio") {
    const checked = form.querySelector(`[name="${name}"]:checked`);
    if (checked && checked.value) formData.set(name, checked.value);
    return;
  }
  const value = String(first.value || "").trim();
  if (value) formData.set(name, value);
}

function appendFields(formData, fieldNames) {
  fieldNames.forEach((name) => appendOptionalField(formData, name));
}

function collectOutlineFromEditor() {
  const rows = Array.from(outlineEditor.querySelectorAll(".outline-item"));
  return rows.map((row, index) => {
    const title = row.querySelector(".outline-title").value.trim() || `第${index + 1}页`;
    const pointsRaw = row.querySelector(".outline-points").value || "";
    const keyPoints = pointsRaw
      .split(/\r?\n/)
      .map((line) => line.trim())
      .filter(Boolean);
    return {
      page: index + 1,
      title,
      key_points: keyPoints,
    };
  });
}

function getCurrentExportMode() {
  if (state.mode === "replica") {
    return form.querySelector('[name="export_mode"]').value;
  }
  return form.querySelector('[name="export_mode_generate"]').value;
}

function getSelectedPages() {
  return Array.from(state.selectedPages).sort((a, b) => a - b);
}

function setButtonsDisabled(disabled) {
  [prepareBtn, replicaBtn, saveEditsBtn, renderBtn, rerenderBtn, editableBtn, editableSelectedBtn].forEach((btn) => {
    btn.disabled = disabled;
  });
}

function setMode(mode) {
  state.mode = mode;
  const isGenerate = mode === "generate";
  generateFields.classList.toggle("hidden", !isGenerate);
  replicaFields.classList.toggle("hidden", isGenerate);
  prepareBtn.classList.toggle("hidden", !isGenerate);
  replicaBtn.classList.toggle("hidden", isGenerate);
}

function renderOutlineEditor(outline = []) {
  const fingerprint = JSON.stringify(outline);
  if (fingerprint === state.outlineFingerprint) return;
  state.outlineFingerprint = fingerprint;
  outlineEditor.innerHTML = "";
  outline.forEach((slide, index) => {
    const item = document.createElement("article");
    item.className = "outline-item";
    item.dataset.page = String(slide.page || index + 1);
    item.innerHTML = `
      <div class="outline-head">
        <span class="outline-index">第 ${index + 1} 页</span>
        <button type="button" class="mini-btn move-up">上移</button>
        <button type="button" class="mini-btn move-down">下移</button>
      </div>
      <label>
        标题
        <input class="outline-title" type="text" value="${escapeHtml(slide.title || "")}" />
      </label>
      <label>
        要点（每行一条）
        <textarea class="outline-points" rows="4">${escapeHtml((slide.key_points || []).join("\n"))}</textarea>
      </label>
    `;
    const up = item.querySelector(".move-up");
    const down = item.querySelector(".move-down");
    up.addEventListener("click", () => moveOutlineItem(item, -1));
    down.addEventListener("click", () => moveOutlineItem(item, 1));
    outlineEditor.appendChild(item);
  });
}

function moveOutlineItem(item, delta) {
  const siblings = Array.from(outlineEditor.children);
  const index = siblings.indexOf(item);
  const target = index + delta;
  if (target < 0 || target >= siblings.length) return;
  if (delta < 0) {
    outlineEditor.insertBefore(item, siblings[target]);
  } else {
    outlineEditor.insertBefore(siblings[target], item);
  }
  state.outlineFingerprint = "";
  renderOutlineIndex();
}

function renderOutlineIndex() {
  Array.from(outlineEditor.querySelectorAll(".outline-item")).forEach((item, idx) => {
    const label = item.querySelector(".outline-index");
    label.textContent = `第 ${idx + 1} 页`;
  });
}

function renderSlides(slides = []) {
  const fingerprint = JSON.stringify(slides.map((s) => [s.page, s.image_url, s.rendered_at || ""]));
  if (fingerprint === state.slidesFingerprint) {
    refreshSelectedCheckboxState();
    return;
  }
  state.slidesFingerprint = fingerprint;
  slidesGrid.innerHTML = "";
  slides.forEach((slide) => {
    const page = Number(slide.page || 0);
    if (!page) return;
    const renderedAt = slide.rendered_at ? encodeURIComponent(slide.rendered_at) : "";
    const imageSrc = slide.image_url ? `${slide.image_url}${renderedAt ? `?t=${renderedAt}` : ""}` : "";
    const card = document.createElement("article");
    card.className = "slide-card";
    const checked = state.selectedPages.has(page) ? "checked" : "";
    card.innerHTML = `
      <img src="${imageSrc}" alt="slide-${page}" loading="lazy" />
      <div class="slide-meta">
        <label class="slide-select">
          <input type="checkbox" data-page="${page}" ${checked} />
          <span>选择第 ${page} 页</span>
        </label>
        <h3>${escapeHtml(slide.title || `第${page}页`)}</h3>
        <details>
          <summary>查看Prompt</summary>
          <pre class="prompt">${escapeHtml(slide.prompt || "")}</pre>
        </details>
      </div>
    `;
    const checkbox = card.querySelector('input[type="checkbox"]');
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.selectedPages.add(page);
      else state.selectedPages.delete(page);
    });
    slidesGrid.appendChild(card);
  });
}

function refreshSelectedCheckboxState() {
  const checkboxes = slidesGrid.querySelectorAll('input[type="checkbox"][data-page]');
  checkboxes.forEach((checkbox) => {
    const page = Number(checkbox.dataset.page || 0);
    checkbox.checked = state.selectedPages.has(page);
  });
}

function renderEditableCards(editableDeck, title = "可编辑结果") {
  if (!editableDeck || !editableDeck.slides || !editableDeck.slides.length) return false;
  const wrapper = document.createElement("section");
  wrapper.className = "editable-group";
  const header = document.createElement("p");
  header.className = "hint";
  header.textContent = title;
  wrapper.appendChild(header);
  editableDeck.slides.forEach((slide) => {
    const page = Number(slide.page || 0);
    const card = document.createElement("article");
    card.className = "editable-card";
    const previewHtmlUrl = slide.preview_html_url || "";
    card.innerHTML = `
      <h3>第 ${page} 页 | 剩余占位符 ${Number(slide.remaining_ph_count || 0)}</h3>
      <div class="editable-links">
        ${previewHtmlUrl ? `<a href="${previewHtmlUrl}" target="_blank" rel="noopener">打开HTML预览</a>` : ""}
        ${slide.preview_pptx_url ? `<a href="${slide.preview_pptx_url}" target="_blank" rel="noopener">下载该页PPT</a>` : ""}
      </div>
      ${previewHtmlUrl ? `<iframe class="editable-frame" loading="lazy" src="${previewHtmlUrl}"></iframe>` : "<p class='hint'>该页暂无可视化预览。</p>"}
    `;
    wrapper.appendChild(card);
  });
  editableGrid.appendChild(wrapper);
  return true;
}

function renderResult(payload) {
  if (!payload) return;
  if (payload.session_id) {
    state.sessionId = payload.session_id;
    sessionIdInput.value = payload.session_id;
  }
  editorCard.classList.remove("hidden");
  resultCard.classList.remove("hidden");
  deckTitleInput.value = payload.deck_title || "";
  stylePromptInput.value = payload.style_prompt || "";
  renderOutlineEditor(payload.outline || []);
  renderSlides(payload.slides || []);

  if (payload.pptx_url) {
    pptxLink.href = payload.pptx_url;
    pptxLink.classList.remove("hidden");
  } else {
    pptxLink.classList.add("hidden");
  }

  const editableDeck = payload.editable_deck && payload.editable_deck.pptx_url ? payload.editable_deck : null;
  if (editableDeck && editableDeck.pptx_url) {
    editableLink.href = editableDeck.pptx_url;
    editableLink.classList.remove("hidden");
  } else {
    editableLink.classList.add("hidden");
  }

  editableGrid.innerHTML = "";
  let hasEditable = false;
  if (editableDeck) {
    hasEditable = renderEditableCards(editableDeck, "完整可编辑PPT预览");
  }
  if (payload.partial_editable_deck) {
    hasEditable = renderEditableCards(payload.partial_editable_deck, "选中页可编辑预览") || hasEditable;
  }
  editablePreviewCard.classList.toggle("hidden", !hasEditable);
}

async function pollJob(jobId) {
  state.pollToken += 1;
  const token = state.pollToken;
  while (token === state.pollToken) {
    const response = await fetch(`/api/generate/status/${jobId}`);
    if (!response.ok) {
      throw new Error(`查询任务状态失败（HTTP ${response.status}）`);
    }
    const job = await response.json();
    updateProgress(job);
    if (job.result_preview) {
      renderResult(job.result_preview);
    }
    if (job.state === "done") {
      if (job.result) renderResult(job.result);
      return;
    }
    if (job.state === "failed") {
      throw new Error(job.error || job.message || "任务失败");
    }
    await sleep(1000);
  }
}

async function runJobRequest(path, formData, startMessage) {
  resetProgress();
  setButtonsDisabled(true);
  setStatus(startMessage);
  try {
    const response = await fetch(path, {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      const errJson = await response.json().catch(() => ({}));
      throw new Error(errJson.detail || `请求失败（HTTP ${response.status}）`);
    }
    const startData = await response.json();
    const jobId = startData.job_id;
    if (startData.session_id) {
      state.sessionId = startData.session_id;
      sessionIdInput.value = startData.session_id;
    }
    if (!jobId) {
      throw new Error("后端未返回 job_id");
    }
    setStatus(`任务已启动（job: ${jobId}）`);
    await pollJob(jobId);
    setStatus("任务完成");
  } catch (error) {
    setStatus(`任务失败：${error.message}`, true);
  } finally {
    setButtonsDisabled(false);
  }
}

async function saveSessionEdits(silent = false) {
  if (!state.sessionId) {
    if (!silent) setStatus("请先创建会话（步骤1）。", true);
    return false;
  }
  const outline = collectOutlineFromEditor();
  if (!outline.length) {
    if (!silent) setStatus("当前大纲为空，无法保存。", true);
    return false;
  }
  const densityNode = form.querySelector('[name="information_density"]');
  const formData = new FormData();
  formData.set("session_id", state.sessionId);
  formData.set("deck_title", deckTitleInput.value.trim());
  const stylePrompt = stylePromptInput.value.trim();
  if (stylePrompt) {
    formData.set("style_prompt", stylePrompt);
  }
  formData.set("information_density", densityNode.value);
  formData.set("outline_json", JSON.stringify(outline));
  try {
    const response = await fetch("/api/workflow/session/update", {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      const errJson = await response.json().catch(() => ({}));
      throw new Error(errJson.detail || `保存失败（HTTP ${response.status}）`);
    }
    const payload = await response.json();
    renderResult(payload);
    if (!silent) setStatus("风格和大纲已保存。");
    return true;
  } catch (error) {
    if (!silent) setStatus(`保存失败：${error.message}`, true);
    return false;
  }
}

function validateStyleMutualExclusive() {
  const styleDesc = String(form.querySelector('[name="style_description"]').value || "").trim();
  const styleFile = form.querySelector('[name="style_template"]').files[0];
  if (styleDesc && styleFile && styleFile.size > 0) {
    setStatus("风格描述与风格图片二选一，请删除其中一个。", true);
    return false;
  }
  return true;
}

async function handlePrepare() {
  if (!validateStyleMutualExclusive()) return;
  const requirement = String(form.querySelector('[name="user_requirement"]').value || "").trim();
  if (!requirement) {
    setStatus("请先填写需求描述。", true);
    return;
  }
  const formData = new FormData();
  formData.set("user_requirement", requirement);
  formData.set("slide_count", String(form.querySelector('[name="slide_count"]').value || "auto").trim() || "auto");
  formData.set("information_density", form.querySelector('[name="information_density"]').value || "medium");

  const styleDesc = String(form.querySelector('[name="style_description"]').value || "").trim();
  if (styleDesc) formData.set("style_description", styleDesc);

  const styleFile = form.querySelector('[name="style_template"]').files[0];
  if (styleFile && styleFile.size > 0) {
    const styleTemplateBase64 = await readFileAsDataUrl(styleFile);
    formData.set("style_template_base64", styleTemplateBase64);
  }

  const sourceFiles = Array.from(form.querySelector('[name="source_files"]').files || []);
  sourceFiles.forEach((file) => formData.append("source_files", file));

  appendFields(formData, MODEL_FIELDS);

  setButtonsDisabled(true);
  setStatus("正在生成风格与大纲，请稍候...");
  try {
    const response = await fetch("/api/workflow/prepare", {
      method: "POST",
      body: formData,
    });
    if (!response.ok) {
      const errJson = await response.json().catch(() => ({}));
      throw new Error(errJson.detail || `准备失败（HTTP ${response.status}）`);
    }
    const payload = await response.json();
    renderResult(payload);
    setStatus("风格与大纲已生成，可先人工编辑确认再渲染。");
  } catch (error) {
    setStatus(`准备失败：${error.message}`, true);
  } finally {
    setButtonsDisabled(false);
  }
}

async function handleRenderAll() {
  const ok = await saveSessionEdits(true);
  if (!ok) return;
  const formData = new FormData();
  formData.set("session_id", state.sessionId);
  formData.set("export_mode", getCurrentExportMode());
  await runJobRequest("/api/workflow/render/start", formData, "正在启动渲染任务...");
}

async function handleRerenderSelected() {
  const pages = getSelectedPages();
  if (!pages.length) {
    setStatus("请先勾选要重绘的页面。", true);
    return;
  }
  const ok = await saveSessionEdits(true);
  if (!ok) return;
  const formData = new FormData();
  formData.set("session_id", state.sessionId);
  formData.set("pages", pages.join(","));
  formData.set("export_mode", getCurrentExportMode());
  await runJobRequest("/api/workflow/slides/regenerate/start", formData, "正在重绘选中页面...");
}

async function handleEditable(fullDeck = true) {
  if (!state.sessionId) {
    setStatus("请先准备并渲染会话。", true);
    return;
  }
  const formData = new FormData();
  formData.set("session_id", state.sessionId);
  if (!fullDeck) {
    const pages = getSelectedPages();
    if (!pages.length) {
      setStatus("请先勾选要生成可编辑预览的页面。", true);
      return;
    }
    formData.set("editable_pages", pages.join(","));
  }
  appendFields(formData, EDITABLE_FIELDS);
  await runJobRequest(
    "/api/workflow/editable/start",
    formData,
    fullDeck ? "正在生成完整可编辑PPT..." : "正在生成选中页可编辑预览..."
  );
}

async function handleReplica() {
  const files = Array.from(form.querySelector('[name="replica_images"]').files || []);
  if (!files.length) {
    setStatus("请先上传复刻图片。", true);
    return;
  }
  const formData = new FormData();
  formData.set("deck_title", String(form.querySelector('[name="replica_deck_title"]').value || "").trim() || "图片复刻结果");
  formData.set("export_mode", form.querySelector('[name="export_mode"]').value || "both");
  formData.set("generate_editable_ppt", form.querySelector('[name="replica_generate_editable"]').value || "false");
  files.forEach((file) => formData.append("replica_images", file));
  appendFields(formData, EDITABLE_FIELDS);
  await runJobRequest("/api/workflow/replica/start", formData, "正在启动图片复刻任务...");
}

async function loadDefaults() {
  try {
    const response = await fetch("/api/workflow/defaults");
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }
    const payload = await response.json();
    defaultsText.textContent = JSON.stringify(payload, null, 2);
  } catch (error) {
    defaultsText.textContent = `加载默认配置失败：${error.message}`;
  }
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

document.querySelectorAll('input[name="workflow_mode"]').forEach((node) => {
  node.addEventListener("change", () => setMode(node.value));
});
prepareBtn.addEventListener("click", handlePrepare);
replicaBtn.addEventListener("click", handleReplica);
saveEditsBtn.addEventListener("click", () => saveSessionEdits(false));
renderBtn.addEventListener("click", handleRenderAll);
rerenderBtn.addEventListener("click", handleRerenderSelected);
editableBtn.addEventListener("click", () => handleEditable(true));
editableSelectedBtn.addEventListener("click", () => handleEditable(false));

if (styleTemplateInput && styleTemplateBase64Input) {
  styleTemplateInput.addEventListener("change", () => {
    styleTemplateBase64Input.value = "";
  });
}

setMode("generate");
loadDefaults();
