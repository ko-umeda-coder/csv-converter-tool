// ============================
// main.js : CSV変換（B2対応・安定版）
// ============================

console.log("✅ main.js 読み込み完了");

const fileInput = document.getElementById('csvFile');
const fileWrapper = document.getElementById('fileWrapper');
const fileName = document.getElementById('fileName');
const convertBtn = document.getElementById('convertBtn');
const downloadBtn = document.getElementById('downloadBtn');
const messageBox = document.getElementById('message');
const courierSelect = document.getElementById('courierSelect');

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
// 宅配会社の選択肢
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
// CSVパース
// ============================
function parseCsv(text) {
  return text
    .trim()
    .split(/\r?\n/)
    .map(line => line.split(",").map(v => v.replace(/^"|"$/g, "").trim()))
    .filter(row => row.length > 1);
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

      if (courier === "yamato") {
        convertedRows = convertToYamato(rows, senderInfo);
      } else {
        showMessage("この運送会社の変換はまだ対応していません。", "error");
        showLoading(false);
        return;
      }

      showMessage("✅ ヤマト運輸用に変換が完了しました。", "success");

      // ダウンロードボタン有効化
      downloadBtn.style.display = "block";
      downloadBtn.disabled = false;
      downloadBtn.className = "btn btn-primary";
    } catch (err) {
      console.error(err);
      showMessage("変換中にエラーが発生しました。", "error");
    } finally {
      showLoading(false);
    }
  });
}

// ============================
// ヤマト運輸用変換
// ============================
function convertToYamato(rows, sender) {
  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  const shipDate = `${yyyy}/${mm}/${dd}`;

  // 1行目がヘッダーならスキップ
  const start = rows[0].some(v => v.includes("col")) ? 1 : 0;

  return rows.slice(start).map(r => [
    r[1] || "",            // col2 お客様管理番号
    "0",                   // 送り状種類
    "0",                   // クール区分
    shipDate,              // 出荷予定日
    r[13] || "",           // col14 お届け先電話番号
    r[10] || "",           // col11 お届け先郵便番号
    r[11] || "",           // col12 お届け先住所１
    r[12] || "",           // col13 お届け先氏名
    "ブーケフレーム加工品", // 品名1
    sender.name,
    sender.phone,
    sender.postal,
    sender.address
  ]);
}

// ============================
// ダウンロードボタン
// ============================
function setupDownloadButton() {
  downloadBtn.addEventListener("click", () => {
    if (convertedRows.length === 0) {
      alert("変換後のデータがありません。");
      return;
    }
    downloadCsv(convertedRows, "yamato_b2_import.csv");
  });
}

// ============================
// CSVダウンロード（Shift_JIS）
// ============================
function downloadCsv(rows, filename) {
  if (typeof Encoding === "undefined") {
    alert("encoding-japaneseが読み込まれていません。index.htmlの順序を確認してください。");
    return;
  }

  const csv = rows.map(r => r.join(",")).join("\r\n");
  const sjisArray = Encoding.convert(Encoding.stringToCode(csv), "SJIS", "UNICODE");
  const blob = new Blob([new Uint8Array(sjisArray)], { type: "text/csv" });

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
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
    overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">変換中...</div></div>`;
    document.body.appendChild(overlay);
  }
  overlay.style.display = show ? "flex" : "none";
}
