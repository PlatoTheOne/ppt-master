const elements = {
  deckTitle: document.getElementById("deckTitle"),
  pageCount: document.getElementById("pageCount"),
  styleKey: document.getElementById("styleKey"),
  exportMode: document.getElementById("exportMode"),
  templateCards: document.getElementById("templateCards"),
  planner: document.getElementById("planner"),
  language: document.getElementById("language"),
  model: document.getElementById("model"),
  baseUrl: document.getElementById("baseUrl"),
  apiKey: document.getElementById("apiKey"),
  fileInput: document.getElementById("fileInput"),
  dropzone: document.getElementById("dropzone"),
  fileMeta: document.getElementById("fileMeta"),
  sourceText: document.getElementById("sourceText"),
  sourceUrl: document.getElementById("sourceUrl"),
  generateBtn: document.getElementById("generateBtn"),
  statusPill: document.getElementById("statusPill"),
  stageLabel: document.getElementById("stageLabel"),
  progressText: document.getElementById("progressText"),
  progressBar: document.getElementById("progressBar"),
  nativeLink: document.getElementById("nativeLink"),
  legacyLink: document.getElementById("legacyLink"),
  webDeckLink: document.getElementById("webDeckLink"),
  projectDir: document.getElementById("projectDir"),
  exportBtn: document.getElementById("exportBtn"),
  refreshDraftBtn: document.getElementById("refreshDraftBtn"),
  retryExportBtn: document.getElementById("retryExportBtn"),
  openProjectBtn: document.getElementById("openProjectBtn"),
  openExportsBtn: document.getElementById("openExportsBtn"),
  editorDeckTitle: document.getElementById("editorDeckTitle"),
  editorDeckSubtitle: document.getElementById("editorDeckSubtitle"),
  slideEditorList: document.getElementById("slideEditorList"),
  previewStage: document.getElementById("previewStage"),
  previewFrame: document.getElementById("previewFrame"),
  previewCounter: document.getElementById("previewCounter"),
  previewCaption: document.getElementById("previewCaption"),
  previewThumbs: document.getElementById("previewThumbs"),
  prevSlideBtn: document.getElementById("prevSlideBtn"),
  nextSlideBtn: document.getElementById("nextSlideBtn"),
  refreshHistoryBtn: document.getElementById("refreshHistoryBtn"),
  historyList: document.getElementById("historyList"),
  logOutput: document.getElementById("logOutput"),
};

let currentJobId = null;
let pollTimer = null;
let uploadedFile = null;
let slideDeck = [];
let activeSlideIndex = 0;
let currentProjectDir = "";
let currentExportsDir = "";
let currentEditor = { deckTitle: "", deckSubtitle: "", slides: [] };
let logLines = [];
let currentExportMode = "pptx";

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function setStatus(status, label) {
  elements.statusPill.className = `status-pill ${status}`;
  elements.statusPill.textContent = label;
}

function setProgress(progress = 0, label = "等待开始") {
  const safeProgress = Math.max(0, Math.min(100, Number(progress) || 0));
  elements.progressBar.style.width = `${safeProgress}%`;
  elements.progressText.textContent = `${safeProgress}%`;
  elements.stageLabel.textContent = label;
}

function setLink(node, href, label) {
  node.textContent = label;
  if (href) {
    node.href = href;
    node.classList.remove("disabled");
  } else {
    node.href = "#";
    node.classList.add("disabled");
  }
}

function setTemplateSelection(styleKey) {
  elements.styleKey.value = styleKey;
  Array.from(elements.templateCards.querySelectorAll(".template-card")).forEach((card) => {
    card.classList.toggle("active", card.dataset.style === styleKey);
  });
}

function syncExportButtonLabel() {
  const mode = elements.exportMode.value;
  if (mode === "web") {
    elements.exportBtn.textContent = "导出 Web Deck";
    return;
  }
  if (mode === "both") {
    elements.exportBtn.textContent = "导出双版本";
    return;
  }
  elements.exportBtn.textContent = "导出 PPTX";
}

function resetResultActions() {
  currentProjectDir = "";
  currentExportsDir = "";
  elements.exportBtn.disabled = true;
  elements.refreshDraftBtn.disabled = true;
  elements.retryExportBtn.disabled = true;
  elements.openProjectBtn.disabled = true;
  elements.openExportsBtn.disabled = true;
}

function clearPreview() {
  slideDeck = [];
  activeSlideIndex = 0;
  elements.previewStage.classList.add("is-empty");
  elements.previewFrame.removeAttribute("src");
  elements.previewCounter.textContent = "0 / 0";
  elements.previewCaption.textContent = "尚未生成页面";
  elements.previewThumbs.innerHTML = '<div class="preview-empty compact">生成完成后，这里会显示缩略图。</div>';
  elements.prevSlideBtn.disabled = true;
  elements.nextSlideBtn.disabled = true;
}

function renderActiveSlide() {
  if (!slideDeck.length) {
    clearPreview();
    return;
  }

  const slide = slideDeck[activeSlideIndex];
  elements.previewStage.classList.remove("is-empty");
  elements.previewFrame.src = slide.src;
  elements.previewCounter.textContent = `${activeSlideIndex + 1} / ${slideDeck.length}`;
  elements.previewCaption.textContent = slide.label;
  elements.prevSlideBtn.disabled = activeSlideIndex === 0;
  elements.nextSlideBtn.disabled = activeSlideIndex === slideDeck.length - 1;

  Array.from(elements.previewThumbs.querySelectorAll(".thumb-btn")).forEach((button, index) => {
    button.classList.toggle("active", index === activeSlideIndex);
  });
}

function renderPreviews(slides = [], labels = []) {
  if (!slides.length) {
    clearPreview();
    return;
  }

  slideDeck = slides.map((src, index) => ({
    src,
    label: labels[index] || `Slide ${String(index + 1).padStart(2, "0")}`,
  }));

  activeSlideIndex = 0;
  elements.previewThumbs.innerHTML = slideDeck
    .map(
      (slide, index) => `
        <button class="thumb-btn ${index === 0 ? "active" : ""}" type="button" data-index="${index}">
          <iframe src="${escapeHtml(slide.src)}" title="${escapeHtml(slide.label)}"></iframe>
          <span>${escapeHtml(slide.label)}</span>
        </button>
      `,
    )
    .join("");

  Array.from(elements.previewThumbs.querySelectorAll(".thumb-btn")).forEach((button) => {
    button.addEventListener("click", () => {
      activeSlideIndex = Number(button.dataset.index || 0);
      renderActiveSlide();
    });
  });

  renderActiveSlide();
}

function appendLogs(logs = [], replace = false) {
  if (replace) {
    logLines = [...logs];
  } else if (logs.length) {
    logLines.push(...logs);
  }
  elements.logOutput.textContent = logLines.length ? logLines.join("\n") : "准备就绪。";
  elements.logOutput.scrollTop = elements.logOutput.scrollHeight;
}

function resetEditor() {
  currentEditor = { deckTitle: "", deckSubtitle: "", slides: [] };
  elements.editorDeckTitle.value = "";
  elements.editorDeckSubtitle.value = "";
  elements.slideEditorList.innerHTML = '<div class="preview-empty">生成草稿后，这里可以调整页面顺序和标题。</div>';
}

function moveSlide(index, direction) {
  const target = index + direction;
  if (target < 0 || target >= currentEditor.slides.length) return;
  const nextSlides = [...currentEditor.slides];
  [nextSlides[index], nextSlides[target]] = [nextSlides[target], nextSlides[index]];
  currentEditor.slides = nextSlides;
  renderEditor();
}

function renderEditor() {
  elements.editorDeckTitle.value = currentEditor.deckTitle || "";
  elements.editorDeckSubtitle.value = currentEditor.deckSubtitle || "";

  if (!currentEditor.slides.length) {
    elements.slideEditorList.innerHTML = '<div class="preview-empty">生成草稿后，这里可以调整页面顺序和标题。</div>';
    return;
  }

  elements.slideEditorList.innerHTML = currentEditor.slides
    .map(
      (slide, index) => `
        <article class="slide-editor-card" data-index="${index}">
          <div class="slide-editor-index">P${String(index + 1).padStart(2, "0")}</div>
          <div class="slide-editor-main">
            <input type="text" value="${escapeHtml(slide.title || "")}" data-role="title">
            <div class="slide-editor-meta">${escapeHtml(slide.type || "slide")}</div>
          </div>
          <div class="slide-editor-actions">
            <button class="mini-btn" type="button" data-action="up">上移</button>
            <button class="mini-btn" type="button" data-action="down">下移</button>
          </div>
        </article>
      `,
    )
    .join("");

  Array.from(elements.slideEditorList.querySelectorAll(".slide-editor-card")).forEach((card) => {
    const index = Number(card.dataset.index || 0);
    card.querySelector('[data-role="title"]').addEventListener("input", (event) => {
      currentEditor.slides[index].title = event.target.value;
    });
    card.querySelector('[data-action="up"]').addEventListener("click", () => moveSlide(index, -1));
    card.querySelector('[data-action="down"]').addEventListener("click", () => moveSlide(index, 1));
  });
}

function hydrateEditor(editor) {
  currentEditor = {
    deckTitle: editor?.deckTitle || "",
    deckSubtitle: editor?.deckSubtitle || "",
    slides: (editor?.slides || []).map((slide) => ({ ...slide })),
  };
  renderEditor();
}

function currentEditorPayload() {
  return {
    deckTitle: elements.editorDeckTitle.value.trim(),
    deckSubtitle: elements.editorDeckSubtitle.value.trim(),
    slides: currentEditor.slides.map((slide) => ({
      id: slide.id,
      title: slide.title,
    })),
  };
}

async function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      resolve(result.includes(",") ? result.split(",")[1] : result);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function buildDraftPayload() {
  const payload = {
    deckTitle: elements.deckTitle.value.trim(),
    pageCount: Number(elements.pageCount.value || 6),
    styleKey: elements.styleKey.value,
    exportMode: elements.exportMode.value,
    planner: elements.planner.value,
    language: elements.language.value,
    model: elements.model.value.trim(),
    baseUrl: elements.baseUrl.value.trim(),
    apiKey: elements.apiKey.value.trim(),
    sourceText: elements.sourceText.value.trim(),
    sourceUrl: elements.sourceUrl.value.trim(),
  };

  if (uploadedFile) {
    payload.fileName = uploadedFile.name;
    payload.fileContentBase64 = await fileToBase64(uploadedFile);
  }

  return payload;
}

async function fetchConfig() {
  const response = await fetch("/api/config");
  const data = await response.json();
  if (!elements.model.value.trim()) {
    elements.model.value = data.defaultModel || "gpt-4.1-mini";
  }
}

async function openRepoPath(path) {
  const response = await fetch("/api/open-path", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ path }),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "打开目录失败");
  }
}

function formatTimestamp(unixSeconds) {
  return new Date(unixSeconds * 1000).toLocaleString();
}

function normalizeSlideLabels(editor) {
  return (editor?.slides || []).map((slide, index) => slide.title || `Slide ${String(index + 1).padStart(2, "0")}`);
}

async function loadHistory() {
  elements.historyList.innerHTML = '<div class="preview-empty">正在读取最近生成的项目...</div>';
  try {
    const response = await fetch("/api/history");
    const data = await response.json();
    const items = data.items || [];

    if (!items.length) {
      elements.historyList.innerHTML = '<div class="preview-empty">这里会显示最近生成的项目。</div>';
      return;
    }

    elements.historyList.innerHTML = items
      .map(
        (item, index) => `
          <article class="history-card">
            <div class="history-preview">
              ${item.slides?.length ? `<iframe src="${escapeHtml(item.slides[0])}" title="${escapeHtml(item.title)}"></iframe>` : '<div class="preview-empty compact">没有预览</div>'}
            </div>
            <div class="history-body">
              <div class="history-title-row">
                <strong>${escapeHtml(item.title)}</strong>
                <small>${escapeHtml(formatTimestamp(item.updatedAt))}</small>
              </div>
              <div class="history-meta">${escapeHtml(item.projectDir)}</div>
              <div class="history-links">
                ${item.nativePptx ? `<a class="file-link" href="${escapeHtml(item.nativePptx)}" target="_blank" rel="noreferrer">下载原生 PPTX</a>` : ""}
                ${item.legacyPptx ? `<a class="file-link" href="${escapeHtml(item.legacyPptx)}" target="_blank" rel="noreferrer">下载 SVG 参考版</a>` : ""}
                ${item.webDeck ? `<a class="file-link" href="${escapeHtml(item.webDeck)}" target="_blank" rel="noreferrer">打开 Web Deck</a>` : ""}
                ${!item.nativePptx && !item.webDeck ? '<span class="history-meta">尚未导出</span>' : ""}
              </div>
              <div class="history-actions">
                <button class="ghost-btn history-open-project" type="button" data-path="${escapeHtml(item.projectDir)}">打开项目目录</button>
                ${item.exportsDir ? `<button class="ghost-btn history-open-exports" type="button" data-path="${escapeHtml(item.exportsDir)}">打开导出目录</button>` : ""}
                ${item.hasDraft ? `<button class="ghost-btn history-load-draft" type="button" data-path="${escapeHtml(item.projectDir)}" data-index="${index}">加载草稿</button>` : ""}
              </div>
            </div>
          </article>
        `,
      )
      .join("");

    Array.from(elements.historyList.querySelectorAll(".history-open-project")).forEach((button) => {
      button.addEventListener("click", async () => {
        try {
          await openRepoPath(button.dataset.path);
        } catch (error) {
          appendLogs([`打开项目目录失败: ${error.message}`]);
        }
      });
    });

    Array.from(elements.historyList.querySelectorAll(".history-open-exports")).forEach((button) => {
      button.addEventListener("click", async () => {
        try {
          await openRepoPath(button.dataset.path);
        } catch (error) {
          appendLogs([`打开导出目录失败: ${error.message}`]);
        }
      });
    });

    Array.from(elements.historyList.querySelectorAll(".history-load-draft")).forEach((button) => {
      button.addEventListener("click", async () => {
        await loadDraft(button.dataset.path);
      });
    });
  } catch (error) {
    elements.historyList.innerHTML = `<div class="preview-empty">读取历史失败: ${escapeHtml(error.message)}</div>`;
  }
}

function applyDraftResult(result) {
  currentProjectDir = result.projectDir || "";
  currentExportsDir = currentProjectDir ? `${currentProjectDir}/exports` : "";
  currentExportMode = result?.sourceMeta?.exportMode || currentExportMode;
  elements.exportMode.value = currentExportMode;
  syncExportButtonLabel();
  elements.projectDir.textContent = currentProjectDir || "-";
  elements.refreshDraftBtn.disabled = !currentProjectDir;
  elements.exportBtn.disabled = !currentProjectDir;
  elements.openProjectBtn.disabled = !currentProjectDir;
  elements.openExportsBtn.disabled = !currentExportsDir;
  hydrateEditor(result.editor || {});
  renderPreviews(result.slides || [], normalizeSlideLabels(result.editor));
}

function applyExportResult(result) {
  currentProjectDir = result.projectDir || "";
  currentExportsDir = currentProjectDir ? `${currentProjectDir}/exports` : "";
  elements.projectDir.textContent = currentProjectDir || "-";
  setLink(elements.nativeLink, result.nativePptx, result.nativePptx ? "下载原生 PPTX" : "本次未导出");
  setLink(elements.legacyLink, result.legacyPptx, result.legacyPptx ? "下载 SVG 参考版" : "本次未导出");
  setLink(elements.webDeckLink, result.webDeck, result.webDeck ? "打开 Web Deck" : "本次未导出");
  elements.retryExportBtn.disabled = !currentProjectDir;
  elements.openProjectBtn.disabled = !currentProjectDir;
  elements.openExportsBtn.disabled = !currentExportsDir;
  hydrateEditor(result.editor || {});
  renderPreviews(result.slides || [], normalizeSlideLabels(result.editor));
}

async function loadDraft(projectDir) {
  try {
    const response = await fetch(`/api/drafts/${encodeURIComponent(projectDir)}`);
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "加载草稿失败");
    }
    setStatus("completed", "草稿已加载");
    setProgress(100, "Draft ready");
    setLink(elements.nativeLink, null, "尚未生成");
    setLink(elements.legacyLink, null, "尚未生成");
    setLink(elements.webDeckLink, null, "尚未生成");
    elements.retryExportBtn.disabled = true;
    applyDraftResult(data);
    appendLogs([`已加载草稿: ${projectDir}`]);
  } catch (error) {
    appendLogs([`加载草稿失败: ${error.message}`]);
  }
}

async function pollJob(jobId) {
  const response = await fetch(`/api/jobs/${jobId}`);
  const data = await response.json();
  appendLogs(data.logs || [], true);
  setProgress(data.progress || 0, data.stageLabel || "Running");

  if (data.status === "queued") {
    setStatus("idle", "排队中");
    return;
  }

  if (data.status === "running") {
    setStatus("running", "进行中");
    return;
  }

  if (data.status === "completed") {
    clearInterval(pollTimer);
    pollTimer = null;
    elements.generateBtn.disabled = false;
    elements.exportBtn.disabled = !currentProjectDir;
    setStatus("completed", "已完成");
    if (data.action === "draft" || data.action === "retry-render") {
      setLink(elements.nativeLink, null, "尚未生成");
      setLink(elements.legacyLink, null, "尚未生成");
      setLink(elements.webDeckLink, null, "尚未生成");
      elements.retryExportBtn.disabled = true;
      applyDraftResult(data.result || {});
    } else {
      applyExportResult(data.result || {});
    }
    loadHistory();
    return;
  }

  if (data.status === "failed") {
    clearInterval(pollTimer);
    pollTimer = null;
    elements.generateBtn.disabled = false;
    elements.exportBtn.disabled = !currentProjectDir;
    setStatus("failed", "失败");
  }
}

async function startJob(url, payload, pendingLabel) {
  if (pollTimer) {
    clearInterval(pollTimer);
    pollTimer = null;
  }
  setStatus("running", pendingLabel);
  setProgress(8, "Queued");
  appendLogs([`任务已提交: ${pendingLabel}`], true);
  const response = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || "任务提交失败");
  }
  currentJobId = data.jobId;
  pollTimer = setInterval(() => pollJob(currentJobId), 1200);
  pollJob(currentJobId);
}

async function createDraftJob() {
  currentExportMode = elements.exportMode.value;
  elements.generateBtn.disabled = true;
  resetResultActions();
  setLink(elements.nativeLink, null, "尚未生成");
  setLink(elements.legacyLink, null, "尚未生成");
  setLink(elements.webDeckLink, null, "尚未生成");
  elements.projectDir.textContent = "-";
  clearPreview();
  resetEditor();
  const payload = await buildDraftPayload();
  await startJob("/api/jobs", payload, "生成草稿");
}

async function exportDraftJob() {
  if (!currentProjectDir) return;
  currentExportMode = elements.exportMode.value;
  elements.exportBtn.disabled = true;
  await startJob(
    "/api/export-jobs",
    {
      projectDir: currentProjectDir,
      exportMode: currentExportMode,
      edits: currentEditorPayload(),
    },
    "导出成品",
  );
}

async function retryJob(step) {
  if (!currentProjectDir) return;
  await startJob(
    "/api/retry-jobs",
    {
      projectDir: currentProjectDir,
      step,
      exportMode: currentExportMode,
    },
    step === "render" ? "重渲染草稿" : "重试导出",
  );
}

elements.fileInput.addEventListener("change", (event) => {
  uploadedFile = event.target.files[0] || null;
  elements.fileMeta.textContent = uploadedFile
    ? `${uploadedFile.name} · ${Math.round(uploadedFile.size / 1024)} KB`
    : "还没有选择文件";
});

elements.styleKey.addEventListener("change", () => {
  setTemplateSelection(elements.styleKey.value);
});

elements.exportMode.addEventListener("change", () => {
  currentExportMode = elements.exportMode.value;
  syncExportButtonLabel();
});

Array.from(elements.templateCards.querySelectorAll(".template-card")).forEach((card) => {
  card.addEventListener("click", () => {
    setTemplateSelection(card.dataset.style);
  });
});

["dragenter", "dragover"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.add("is-dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  elements.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    elements.dropzone.classList.remove("is-dragover");
  });
});

elements.dropzone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files || [];
  if (!file) return;
  uploadedFile = file;
  const transfer = new DataTransfer();
  transfer.items.add(file);
  elements.fileInput.files = transfer.files;
  elements.fileMeta.textContent = `${uploadedFile.name} · ${Math.round(uploadedFile.size / 1024)} KB`;
});

elements.prevSlideBtn.addEventListener("click", () => {
  if (activeSlideIndex > 0) {
    activeSlideIndex -= 1;
    renderActiveSlide();
  }
});

elements.nextSlideBtn.addEventListener("click", () => {
  if (activeSlideIndex < slideDeck.length - 1) {
    activeSlideIndex += 1;
    renderActiveSlide();
  }
});

elements.generateBtn.addEventListener("click", async () => {
  try {
    await createDraftJob();
  } catch (error) {
    setStatus("failed", "失败");
    elements.generateBtn.disabled = false;
    appendLogs([`生成草稿失败: ${error.message}`]);
  }
});

elements.exportBtn.addEventListener("click", async () => {
  try {
    await exportDraftJob();
  } catch (error) {
    setStatus("failed", "失败");
    elements.exportBtn.disabled = false;
    appendLogs([`导出失败: ${error.message}`]);
  }
});

elements.refreshDraftBtn.addEventListener("click", async () => {
  try {
    await retryJob("render");
  } catch (error) {
    appendLogs([`重渲染失败: ${error.message}`]);
  }
});

elements.retryExportBtn.addEventListener("click", async () => {
  try {
    await retryJob("export");
  } catch (error) {
    appendLogs([`重试导出失败: ${error.message}`]);
  }
});

elements.openProjectBtn.addEventListener("click", async () => {
  if (!currentProjectDir) return;
  try {
    await openRepoPath(currentProjectDir);
  } catch (error) {
    appendLogs([`打开项目目录失败: ${error.message}`]);
  }
});

elements.openExportsBtn.addEventListener("click", async () => {
  if (!currentExportsDir) return;
  try {
    await openRepoPath(currentExportsDir);
  } catch (error) {
    appendLogs([`打开导出目录失败: ${error.message}`]);
  }
});

elements.refreshHistoryBtn.addEventListener("click", () => {
  loadHistory();
});

fetchConfig();
setTemplateSelection(elements.styleKey.value);
setStatus("idle", "待命");
setProgress(0, "等待开始");
syncExportButtonLabel();
appendLogs([], true);
clearPreview();
resetEditor();
resetResultActions();
loadHistory();
