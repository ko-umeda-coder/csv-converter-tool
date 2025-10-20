// ============================
// main.js : ヤマトテンプレート自動転記＋Excel出力版
// ============================

console.log("✅ main.js 読み込み完了");

const fileInput = document.getElementById("csvFile");
const fileWrapper = document.getElementById("fileWrapper");
const fileName = document.getElementById("fileName");
const convertBtn = document.getElementById("convertBtn");
const downloadBtn = document.getElementById("downloadBtn");
const messageBox = document.getElementById("message");
const courierSelect = document.getElementById("courierSelect");

let mergedWorkbook = null;
let originalFileName = "";

// ============================
// 初期化
// ============================
setupCourierOptions();
setupFileInput();
setupConvertButton();
setupDownloadButton();

// ============================
// 宅配会社選択
// ============================
function setupCourierOptions() {
  const options = [
    { value: "yamato", text: "ヤマト運輸" },
    { value: "sagawa", text: "佐川急便（今後対応予定）" },
    { value: "japanpost", text: "日本郵政（今後対応予定）" },
  ];
  courierSelect.innerHTML = options
    .map((o) => `<option value="${o.value}">${o.text}</option>`)
    .join("");
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
    overlay.innerHTML = `
      <div class="loading-content">
        <div class="spinner"></div>
        <div class="loading-text">処理中です...</div>
      </div>`;
    document.body.appendChild(overlay);
  }
  overlay.style.display = show ? "flex" : "none";
}

// ============================
// 送り主情報取得
// ============================
function getSenderInfo() {
  return {
    name: document.getElementById("senderName").value.trim(),
    postal: document.getElementById("senderPostal").value.trim(),
    address: document.getElementById("senderAddress").value.trim(),
    phone: document.getElementById("senderPhone").value.trim(),
  };
}

// ============================
// CSV→テンプレート転記
// ============================
async function mergeToYamatoTemplate(csvFile, templateUrl, sender) {
  // CSV読み込み
  const csvText = await csvFile.text();
  const rows = csvText
    .trim()
    .split(/\r?\n/)
    .map((line) => line.split(","))
    .filter((r) => r.length > 1);

  // テンプレート読込
  const response = await fetch(templateUrl);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets["外部データ取り込み基本レイアウト"];

  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, "0");
  const dd = String(today.getDate()).padStart(2, "0");
  const shipDate = `${yyyy}/${mm}/${dd}`;

  // 1行目ヘッダーの場合スキップ
  const start = rows[0].some((v) => v.includes("col")) ? 1 : 0;
  let rowExcel = 2; // B2テンプレでは2行目から

  rows.slice(start).forEach((r) => {
    sheet[`A${rowExcel}`] = { v: r[1] || "" }; // お客様管理番号 col2
    sheet[`B${rowExcel}`] = { v: "0" }; // 送り状種類
    sheet[`C${rowExcel}`] = { v: "0" }; // クール区分
    sheet[`E${rowExcel}`] = { v: shipDate }; // 出荷予定日
    sheet[`I${rowExcel}`] = { v: r[13] || "" }; // 電話番号 col14
    sheet[`K${rowExcel}`] = { v: r[10] || "" }; // 郵便番号 col11
    sheet[`L${rowExcel}`] = { v: r[11] || "" }; // 住所１ col12
    sheet[`M${rowExcel}`] = { v: r[12] || "" }; // 氏名 col13
    sheet[`U${rowExcel}`] = { v: "ブーケフレーム加工品" }; // 品名1
    sheet[`BF${rowExcel}`] = { v: sender.name }; // ご依頼主
    sheet[`BG${rowExcel}`] = { v: sender.phone }; // ご依頼主電話番号
    sheet[`BI${rowExcel}`] = { v: sender.postal }; // ご依頼主郵便番号
    sheet[`BJ${rowExcel}`] = { v: sender.address }; // ご依頼主住所
    rowExcel++;
  });

  return workbook;
}

// ============================
// 変換ボタン
// ============================
function setupConvertButton() {
  convertBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    const courier = courierSelect.value;
    if (!file || courier !== "yamato") {
      showMessage("ヤマト運輸のみ対応しています。", "error");
      return;
    }

    showLoading(true);
    showMessage("ヤマト運輸テンプレートに転記中...", "info");

    try {
      const sender = getSenderInfo();
      const templatePath = "./js/newb2web_template1.xlsx"; // GitHub Pagesの配置場所
      mergedWorkbook = await mergeToYamatoTemplate(file, templatePath, sender);

      showMessage("✅ 転記が完了しました。ダウンロードできます。", "success");
      downloadBtn.style.display = "block";
      downloadBtn.disabled = false;
      downloadBtn.className = "btn btn-primary";
    } catch (err) {
      console.error(err);
      showMessage("転記中にエラーが発生しました。", "error");
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
    if (!mergedWorkbook) {
      alert("変換後データがありません。");
      return;
    }
    XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
  });
}
