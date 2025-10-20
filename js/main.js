// ============================
// main.js : CSV変換メイン処理
// ============================

console.log("✅ main.js 読み込み完了");

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
// 初期化（DOMContentLoaded問題の修正版）
// ----------------------------
setupFileInput();
setupConvertButton();
setupDownloadButton();

// ----------------------------
// ファイル選択イベント
// ----------------------------
function setupFileInput() {
  console.log("📂 setupFileInput() 実行中");

  // ファイル選択時
  fileInput.addEventListener("change", () => {
    console.log("✅ CSVファイルが選択されました:", fileInput.files);
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
      const rows = parseCsv(text);

      const senderInfo = getSenderInfo();
      if (!validateSenderInfo()) {
        showMessage("送り主情報が正しくありません。郵便番号7桁、電話番号9〜11桁を確認してください。", "error");
        showLoading(false);
        return;
      }

      const format = formats[courier];
      const converted = convertToCourierFormat(rows, senderInfo, format, courier);

      convertedRows = converted;
      showPreview(converted);
      showStats(rows.length - 1, converted.length - 1);
      showMessage("変換が完了しました！", "success");
      downloadBtn.style.display = "block";
    } catch (err) {
      console.error(err);
      showMessage("変換中にエラーが発生しました。", "error");
    } finally {
      showLoading(false);
    }
  });
}

// ----------------------------
// ダウンロードボタン押下
// ----------------------------
function setupDownloadButton() {
  downloadBtn.addEventListener("click", () => {
    if (convertedRows.length === 0) return;
    const courier = courierSelect.value;
    const filename = originalFileName.replace(/\.csv$/, `_${courier}.csv`);
    downloadCsv(convertedRows, filename);
  });
}

// ----------------------------
// 送り主情報の取得
// ----------------------------
function getSenderInfo() {
  return {
    name: document.getElementById("senderName").value.trim(),
    postal: document.getElementById("senderPostal").value.trim(),
    address: document.getElementById("senderAddress").value.trim(),
    phone: document.getElementById("senderPhone").value.trim()
  };
}

// ----------------------------
// メッセージ表示
// ----------------------------
function showMessage(text, type = "info") {
  messageBox.style.display = "block";
  messageBox.textContent = text;
  messageBox.className = "message";
  if (type === "error") messageBox.classList.add("error");
  if (type === "success") messageBox.classList.add("success");
  if (type === "info") {
    messageBox.style.background = "#e2e3e5";
    messageBox.style.borderColor = "#bfc0c1";
    messageBox.style.color = "#383d41";
  }
}

// ----------------------------
// ローディング表示
// ----------------------------
function showLoading(show) {
  let overlay = document.getElementById("loading");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.id = "loading";
    overlay.className = "loading-overlay";
    overlay.innerHTML = `
      <div class="loading-content">
        <div class="spinner"></div>
        <div class="loading-text">変換中です...</div>
      </div>`;
    document.body.appendChild(overlay);
  }
  overlay.style.display = show ? "flex" : "none";
}

// ----------------------------
// 宅配会社フォーマット変換処理
// ----------------------------
function convertToCourierFormat(rows, sender, format, courier) {
  const header = format.columns.map(col => col.header);
  const result = [header];

  const headerRow = rows[0];
  const headerMap = {};
  headerRow.forEach((h, i) => { headerMap[h.trim()] = i; });

  for (let i = 1; i < rows.length; i++) {
    const original = rows[i];
    const newRow = format.columns.map(col => {
      if (col.source?.startsWith("col")) {
        const idx = parseInt(col.source.replace("col", "")) - 1;
        return original[idx] || "";
      } else if (headerMap[col.source] !== undefined) {
        return original[headerMap[col.source]] || "";
      } else if (col.value) {
        return col.value;
      } else if (col.source?.startsWith("sender")) {
        const key = col.source.replace("sender", "").toLowerCase();
        return sender[key] || "";
      } else {
        return "";
      }
    });
    result.push(newRow);
  }

  return result;
}

// ----------------------------
// プレビュー表示
// ----------------------------
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

// ----------------------------
// 統計情報表示
// ----------------------------
function showStats(originalCount, convertedCount) {
  statsBox.innerHTML = `
    <div class="stat-item">
      <div class="stat-number">${originalCount}</div>
      <div class="stat-label">元の件数</div>
    </div>
    <div class="stat-item">
      <div class="stat-number">${convertedCount}</div>
      <div class="stat-label">変換後件数</div>
    </div>
    <div class="stat-item">
      <div class="stat-number">${Object.keys(formats).length}</div>
      <div class="stat-label">対応運送会社</div>
    </div>
  `;
}
