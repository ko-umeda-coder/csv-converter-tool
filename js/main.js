// ============================
// main.js : CSV変換メイン処理（最終安定版）
// ============================

console.log("✅ main.js 読み込み完了");

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

// ============================
// 初期化
// ============================
setupFileInput();
setupConvertButton();
setupDownloadButton();

// ============================
// ファイル選択
// ============================
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
}

// ============================
// 変換ボタン
// ============================
function setupConvertButton() {
  convertBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    const courier = courierSelect.value;
    if (!file || !courier) return;

    showMessage("変換中です...", "info");
    showLoading(true);

    try {
      const text = await file.text();
      const rows = parseCsv(text);
      const senderInfo = getSenderInfo();

      const format = formats[courier];
      const converted = convertToCourierFormat(rows, senderInfo, format, courier);

      convertedRows = converted;
      showPreview(converted);
      showStats(rows.length - 1, converted.length - 1);

      showMessage("変換が完了しました！", "success");

      // ✅ ダウンロードボタンを青く活性化
      downloadBtn.style.display = "block";
      downloadBtn.disabled = false;
      downloadBtn.classList.remove("btn-secondary");
      downloadBtn.classList.add("active");

    } catch (err) {
      console.error(err);
      showMessage("変換中にエラーが発生しました。", "error");
    } finally {
      showLoading(false);
    }
  });
}

// ============================
// ダウンロードボタン
// ============================
function setupDownloadButton() {
  downloadBtn.addEventListener("click", () => {
    if (convertedRows.length === 0) return;
    const courier = courierSelect.value;
    const filename = originalFileName.replace(/\.csv$/, `_${courier}.csv`);
    downloadCsv(convertedRows, filename);
  });
}

// ============================
// 送り主情報取得
// ============================
function getSenderInfo() {
  return {
    name: document.getElementById("senderName").value.trim(),
    postal: document.getElementById("senderPostal").value.trim(),
    address: document.getElementById("senderAddress").value.trim(),
    phone: document.getElementById("senderPhone").value.trim()
  };
}

// ============================
// メッセージ表示
// ============================
function showMessage(text, type = "info") {
  messageBox.style.display = "block";
  messageBox.textContent = text;
  messageBox.className = "message " + type;
}

// ============================
// ローディング表示
// ============================
function showLoading(show) {
  let overlay = document.getElementById("loading");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "loading";
    overlay.className = "loading-overlay";
    overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">変換中です...</div></div>`;
    document.body.appendChild(overlay);
  }
  overlay.style.display = show ? "flex" : "none";
}

// ============================
// プレビュー表示
// ============================
function showPreview(rows) {
  previewSection.style.display = "block";
  const previewRows = rows.slice(0, 6);
  let html = "<table class='table-preview'>";
  previewRows.forEach((r) => {
    html += "<tr>" + r.map(v => `<td>${v}</td>`).join("") + "</tr>";
  });
  html += "</table>";
  previewContent.innerHTML = html;
}

// ============================
// 統計表示
// ============================
function showStats(originalCount, convertedCount) {
  statsBox.innerHTML = `
    <div class="stat-item"><div class="stat-number">${originalCount}</div><div class="stat-label">元の件数</div></div>
    <div class="stat-item"><div class="stat-number">${convertedCount}</div><div class="stat-label">変換後件数</div></div>
  `;
}
