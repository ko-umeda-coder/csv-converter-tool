console.log("✅ main.js 読み込み完了");
// ============================
// main.js : CSV変換メイン処理
// ============================

// ----------------------------
// グローバル変数
// ----------------------------
const fileInput = document.getElementById('csvFile');
const fileWrapper = document.getElementById('fileWrapper');
const fileName = document.getElementById('fileName');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const messageBox = document.getElementById('message');
const courierSelect = document.getElementById('courierSelect');
const previewSection = document.getElementById('previewSection');
const previewContent = document.getElementById('previewContent');
const statsBox = document.getElementById('statsBox');

let convertedRows = [];
let originalFileName = "";

// ----------------------------
// 初期化
// ----------------------------
document.addEventListener("DOMContentLoaded", () => {
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();
});

// ----------------------------
// ファイル選択イベント
// ----------------------------
function setupFileInput() {
  fileInput.addEventListener("change", () => {
    if (fileInput.files.length > 0) {
      const file = fileInput.files[0];
      originalFileName = file.name;
      fileName.textContent = file.name;
      fileWrapper.classList.add("has-file");
      convertBtn.disabled = false;
    } else {
      fileName.textContent = "";
      fileWrapper.classList.remove("has-file");
      convertBtn.disabled = true;
    }
  });

  // ドラッグ＆ドロップ対応
  fileWrapper.addEventListener("dragover", (e) => {
    e.preventDefault();
    fileWrapper.style.borderColor = "var(--primary)";
  });
  fileWrapper.addEventListener("dragleave", () => {
    fileWrapper.style.borderColor = "var(--border)";
  });
  fileWrapper.addEventListener("drop", (e) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".csv")) {
      fileInput.files = e.dataTransfer.files;
      const event = new Event("change");
      fileInput.dispatchEvent(event);
    }
  });
}

// ----------------------------
// 変換ボタン押下イベント
// ----------------------------
function setupConvertButton() {
  convertBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    const courier = courierSelect.value;
    if (!file || !courier) return;

    showMessage("変換中です...", "info");
    showLoading(true);

    try {
      const text = await file.text();
