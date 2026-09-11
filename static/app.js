const csrfToken = document.querySelector('meta[name="csrf-token"]').content;
const fileTab = document.querySelector("#file-tab");
const textTab = document.querySelector("#text-tab");
const filePanel = document.querySelector("#file-panel");
const textPanel = document.querySelector("#text-panel");
const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const fileCard = document.querySelector("#file-card");
const fileName = document.querySelector("#file-name");
const fileSize = document.querySelector("#file-size");
const removeFile = document.querySelector("#remove-file");
const textInput = document.querySelector("#text-input");
const charCount = document.querySelector("#char-count");
const convertButton = document.querySelector("#convert-button");
const statusMessage = document.querySelector("#status-message");
let activeMode = "text";
let selectedFile = null;

function selectMode(mode) {
  activeMode = mode;
  const fileActive = mode === "file";
  fileTab.classList.toggle("active", fileActive);
  textTab.classList.toggle("active", !fileActive);
  fileTab.setAttribute("aria-selected", String(fileActive));
  textTab.setAttribute("aria-selected", String(!fileActive));
  filePanel.hidden = !fileActive;
  textPanel.hidden = fileActive;
  clearStatus();
}

function setFile(file) {
  if (!file) return;
  const extension = file.name.split(".").pop().toLowerCase();
  if (!["txt", "log"].includes(extension)) {
    showStatus("请选择 TXT 或 LOG 文本文件", "error");
    return;
  }
  if (file.size > 2 * 1024 * 1024) {
    showStatus("文件不能超过 2 MB", "error");
    return;
  }
  selectedFile = file;
  fileName.textContent = file.name;
  fileSize.textContent = formatBytes(file.size);
  dropZone.hidden = true;
  fileCard.hidden = false;
  clearStatus();
}

function clearFile() {
  selectedFile = null;
  fileInput.value = "";
  dropZone.hidden = false;
  fileCard.hidden = true;
  clearStatus();
}

function showStatus(message, type = "success") {
  statusMessage.textContent = message;
  statusMessage.className = `message ${type}`;
  statusMessage.hidden = false;
}

function clearStatus() {
  statusMessage.hidden = true;
  statusMessage.textContent = "";
}

function formatBytes(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  return `${(bytes / 1024).toFixed(1)} KB`;
}

function responseFilename(response) {
  const disposition = response.headers.get("Content-Disposition") || "";
  const utf8Match = disposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match) return decodeURIComponent(utf8Match[1]);
  const plainMatch = disposition.match(/filename="?([^";]+)"?/i);
  return plainMatch ? plainMatch[1] : "域名整理.xlsx";
}

async function convert() {
  clearStatus();
  if (activeMode === "file" && !selectedFile) {
    showStatus("请先选择一个文本文件", "error");
    return;
  }
  if (activeMode === "text" && !textInput.value.trim()) {
    showStatus("请先粘贴需要整理的内容", "error");
    textInput.focus();
    return;
  }

  const formData = new FormData();
  formData.append("title", document.querySelector("#title").value.trim());
  if (activeMode === "file") formData.append("file", selectedFile);
  else formData.append("text", textInput.value);

  convertButton.disabled = true;
  convertButton.querySelector("span").textContent = "正在整理…";

  try {
    const response = await fetch("/api/convert", {
      method: "POST",
      headers: { "X-CSRF-Token": csrfToken },
      body: formData,
    });
    if (response.status === 401) {
      window.location.replace("/login");
      return;
    }
    if (!response.ok) {
      const result = await response.json().catch(() => ({}));
      throw new Error(result.error || "生成失败，请稍后重试");
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = responseFilename(response);
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);

    const count = response.headers.get("X-Record-Count");
    showStatus(`整理完成，共识别 ${count || ""} 条域名，文件已开始下载。`);
  } catch (error) {
    showStatus(error.message, "error");
  } finally {
    convertButton.disabled = false;
    convertButton.querySelector("span").textContent = "生成并下载 Excel";
  }
}

fileTab.addEventListener("click", () => selectMode("file"));
textTab.addEventListener("click", () => selectMode("text"));
fileInput.addEventListener("change", () => setFile(fileInput.files[0]));
removeFile.addEventListener("click", clearFile);
textInput.addEventListener("input", () => {
  charCount.textContent = `${textInput.value.length.toLocaleString()} 个字符`;
  clearStatus();
});

["dragenter", "dragover"].forEach((name) => {
  dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    dropZone.classList.add("dragging");
  });
});
["dragleave", "drop"].forEach((name) => {
  dropZone.addEventListener(name, (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
  });
});
dropZone.addEventListener("drop", (event) => setFile(event.dataTransfer.files[0]));
convertButton.addEventListener("click", convert);

document.querySelector("#logout-button").addEventListener("click", async () => {
  const response = await fetch("/api/logout", {
    method: "POST",
    headers: { "X-CSRF-Token": csrfToken },
  });
  if (response.ok || response.status === 401) window.location.replace("/login");
});
