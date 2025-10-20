// ============================
// main.js : CSV変換メイン処理（宅配会社拡張対応版）
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
setupCourierOptions();
setupFileInput();
setupConvertButton();
setupDownloadButton();

// ============================
// 宅配会社の選択肢を追加
// ============================
function setupCourierOptions() {
  const options = [
    { value: "yamato", text: "ヤマト運輸" },
    { value: "sagawa", text: "佐川急便（今後対応予定）" },
    { value: "japanpost", text: "日本郵政（今後対応予定）" }
  ];
  courierSelect.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join("");
}

// ============================
// 各社フォーマット定義
// ============================
const formats = {
  yamato: {
    name: "ヤマト運輸",
    headers: [
      "お客様管理番号", "送り状種類", "クール区分", "出荷予定日",
      "お届け先電話番号", "お届け先郵便番号", "お届け先住所１", "お届け先氏名",
      "品名1", "ご依頼主", "ご依頼主電話番号", "ご依頼主郵便番号", "ご依頼主住所"
    ],
    map: (row, senderInfo) => {
      const today = new Date();
      const yyyy = today.getFullYear();
      const mm = String(today.getMonth() + 1).padStart(2, '0');
      const dd = String(today.getDate()).padStart(2, '0');
      const shipDate = `${yyyy}/${mm}/${dd}`;

      return [
        row.col2 || "",
        "0",
        "0",
        shipDate,
        row.col14 || "",
        row.col11 || "",
        row.col12 || "",
        row.col13 || "",
        "ブーケフレーム加工品",
        senderInfo.name,
        senderInfo.phone,
        senderInfo.postal,
        senderInfo.address
      ];
    }
  },

  sagawa: {
    name: "佐川急便",
    headers: ["（今後追加予定）"],
    map: () => []
  },

  japanpost: {
    name: "日本郵政",
    headers: ["（今後追加予定）"],
    map: () => []
  }
};

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
// CSVパース（簡易版）
// ============================
function parseCsv(text) {
  const rows = text.split(/\r?\n/).filter(r => r.trim() !== "").map(r => r.split(","));
  const headers = rows[0];
  return rows.slice(1).map(r => {
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = r[i] || ""; });
    return obj;
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
      showStats(rows.length, converted.length);

      showMessage(`${format.name}形式に変換が完了しました！`, "success");

      // ダウンロードボタン活性化
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
// 宅配会社別変換処理
// ============================
function convertToCourierFormat(rows, senderInfo, format, courier) {
  if (!format) return [];
  const mapped = rows.map(r => format.map(r, senderInfo));
  return mapped;
}

// ============================
// ダウンロード処理（Shift_JIS）
// ============================
function downloadCsv(rows, filename) {
  if (rows.length === 0) return;

  const csv = rows.map(r => r.join(",")).join("\r\n");
  const sjisArray = Encoding.convert(Encoding.stringToCode(csv), 'SJIS', 'UNICODE');
  const sjisBlob = new Blob([new Uint8Array(sjisArray)], { type: 'text/csv' });

  const url = URL.createObjectURL(sjisBlob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename || "yamato_b2_import.csv";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ============================
// ダウンロードボタン
// ============================
function setupDownloadButton() {
  downloadBtn.addEventListener("click", () => {
    if (convertedRows.length === 0) return;
    const courier = courierSelect.value;
    const filename = courier === "yamato" ? "yamato_b2_import.csv" : `${courier}_export.csv`;
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
  const previewRows = rows.slice(0, 5);
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
