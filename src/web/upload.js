const uploadSection = document.getElementById("uploadSection");
const loadingSection = document.getElementById("loadingSection");
const resultSection = document.getElementById("resultSection");
const errorSection = document.getElementById("errorSection");

const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const uploadButton = document.getElementById("uploadButton");

const loadingFileName = document.getElementById("loadingFileName");
const progressFill = document.getElementById("progressFill");

const previewFrame = document.getElementById("previewFrame");
const downloadLink = document.getElementById("downloadLink");
const openLink = document.getElementById("openLink");
const statsText = document.getElementById("statsText");
const warningsList = document.getElementById("warningsList");

const errorText = document.getElementById("errorText");
const tryAgainButton = document.getElementById("tryAgainButton");
const convertAnotherButton = document.getElementById("convertAnotherButton");

const warningsBox = document.getElementById("warningsBox");

let selectedFile = null;

function showOnly(section) {
  uploadSection.classList.add("hidden");
  loadingSection.classList.add("hidden");
  resultSection.classList.add("hidden");
  errorSection.classList.add("hidden");

  section.classList.remove("hidden");
}

function setProgress(percent) {
  progressFill.style.width = `${percent}%`;
}

function renderWarnings(warnings) {
  warningsList.innerHTML = "";

  if (!warnings || warnings.length === 0) {
    warningsBox.classList.add("hidden");
    return;
  }

  warningsBox.classList.remove("hidden");

  warnings.forEach((warning) => {
    const li = document.createElement("li");
    li.textContent = warning;
    warningsList.appendChild(li);
  });
}

function setSelectedFile(file) {
  selectedFile = file;
  fileInput.files = createFileList(file);

  const title = dropZone.querySelector("h2");
  title.textContent = file.name;
}

function createFileList(file) {
  const dataTransfer = new DataTransfer();
  dataTransfer.items.add(file);
  return dataTransfer.files;
}

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];

  if (file) {
    selectedFile = file;
    const title = dropZone.querySelector("h2");
    title.textContent = file.name;
  }
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragover");

  const file = event.dataTransfer.files[0];

  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    errorText.textContent = "Please upload a .pptx file.";
    showOnly(errorSection);
    return;
  }

  setSelectedFile(file);
});

uploadButton.addEventListener("click", async () => {
  const file = selectedFile || fileInput.files[0];

  if (!file) {
    errorText.textContent = "Please select a PPTX file first.";
    showOnly(errorSection);
    return;
  }

  if (!file.name.toLowerCase().endsWith(".pptx")) {
    errorText.textContent = "Only .pptx files are supported.";
    showOnly(errorSection);
    return;
  }

  const formData = new FormData();
  formData.append("file", file);

  loadingFileName.textContent = `${file.name} · PPTX → reveal.js`;
  setProgress(0);
showOnly(loadingSection);

let fakeProgress = 0;

const progressTimer = setInterval(() => {
  fakeProgress += 12;

  if (fakeProgress < 90) {
    setProgress(fakeProgress);
  }
}, 300);

try {
  const response = await fetch("/api/v1/convert", {
    method: "POST",
    body: formData
  });

  clearInterval(progressTimer);
  setProgress(95);

    const data = await response.json();

    if (!response.ok) {
      const error = new Error(data.message || "Conversion failed.");
      error.code = data.error_code;
      throw error;
    }

    setProgress(100);

    setTimeout(() => {
      previewFrame.src = data.preview_url;
      downloadLink.href = data.download_url;
      openLink.href = data.preview_url;

      statsText.textContent =
        `Slides converted: ${data.statistics?.slide_count ?? 0}. ` +
        `Warnings: ${data.warnings?.length ?? 0}.`;

      renderWarnings(data.warnings);
      showOnly(resultSection);
    }, 300);
    
  } catch (err) {
    clearInterval(progressTimer);
    setProgress(0);
  
    const friendlyMessages = {
      NO_FILE: "Please select a PPTX file before converting.",
      INVALID_EXTENSION: "Only .pptx files are supported.",
      INVALID_PPTX: "This file is not a valid PowerPoint presentation.",
      RESULT_NOT_FOUND: "The conversion result expired. Please convert the file again.",
      MEDIA_NOT_FOUND: "One of the media files could not be loaded.",
      SANITIZER_REJECTED: "This file contains unsafe content and cannot be converted."
    };
  
    errorText.textContent =
      friendlyMessages[err.code] ||
      err.message ||
      "Something went wrong during conversion.";
  
    showOnly(errorSection);
  }
});

convertAnotherButton.addEventListener("click", () => {
  selectedFile = null;
  fileInput.value = "";
  previewFrame.src = "";
  downloadLink.href = "#";
  openLink.href = "#";
  dropZone.querySelector("h2").textContent = "Drop your PPTX file here";
  setProgress(0);
  showOnly(uploadSection);
});

tryAgainButton.addEventListener("click", () => {
  showOnly(uploadSection);
});
