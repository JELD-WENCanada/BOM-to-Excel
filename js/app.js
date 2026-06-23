import { parseBomPdf } from "./parser.js";
import { downloadExcel } from "./exporter.js";

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const errorEl = document.getElementById("error");
const fileNameEl = document.getElementById("file-name");
const itemCountEl = document.getElementById("item-count");
const downloadBtn = document.getElementById("download-btn");
const chooseBtn = document.getElementById("choose-btn");

let currentBom = null;
let currentFileName = "bom.xlsx";

function setStatus(message) {
  statusEl.textContent = message;
  statusEl.hidden = !message;
}

function showError(message) {
  errorEl.textContent = message;
  errorEl.hidden = !message;
}

function resetUi() {
  resultEl.hidden = true;
  showError("");
  setStatus("");
  currentBom = null;
}

function outputFileName(pdfName) {
  const base = pdfName.replace(/\.pdf$/i, "") || "bom";
  return `${base}.xlsx`;
}

async function processFile(file) {
  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".pdf")) {
    showError("Please upload a PDF file.");
    return;
  }

  resetUi();
  setStatus("Reading PDF…");
  resultEl.hidden = true;

  try {
    const bom = await parseBomPdf(file, window.pdfjsLib);
    currentBom = bom;
    currentFileName = outputFileName(file.name);

    fileNameEl.textContent = file.name;
    itemCountEl.textContent = `${bom.itemCount} line items across ${bom.sections.length} categories`;
    resultEl.hidden = false;
    setStatus("");
    showError("");
  } catch (error) {
    console.error(error);
    showError(error.message || "Failed to convert this PDF.");
    setStatus("");
  }
}

downloadBtn.addEventListener("click", async () => {
  if (!currentBom) return;

  try {
    setStatus("Building Excel file…");
    downloadBtn.disabled = true;
    await downloadExcel(currentBom, window.ExcelJS, currentFileName);
    setStatus("Download started.");
  } catch (error) {
    console.error(error);
    showError(error.message || "Failed to create Excel file.");
    setStatus("");
  } finally {
    downloadBtn.disabled = false;
  }
});

chooseBtn.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  processFile(file);
  event.target.value = "";
});

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("drag-over");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("drag-over");
  });
});

dropZone.addEventListener("drop", (event) => {
  const file = event.dataTransfer.files[0];
  processFile(file);
});

dropZone.addEventListener("click", (event) => {
  if (event.target === chooseBtn || chooseBtn.contains(event.target)) return;
  fileInput.click();
});
