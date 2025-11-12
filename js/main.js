// ============================
// XLSXライブラリ読み込み待機
// ============================
const waitForXLSX = () => new Promise(resolve => {
  const check = () => {
    if (window.XLSX) {
      console.log("✅ XLSXライブラリ検出完了");
      resolve();
    } else {
      setTimeout(check, 100);
    }
  };
  check();
});

// ============================
// main.js 本体
// ============================
(async () => {
  await waitForXLSX();
  console.log("✅ main.js 起動");

  const fileInput = document.getElementById("csvFile");
  const fileWrapper = document.getElementById("fileWrapper");
  const fileName = document.getElementById("fileName");
  const convertBtn = document.getElementById("convertBtn");
  const downloadBtn = document.getElementById("downloadBtn");
  const messageBox = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;
  let convertedCSV = null;

  // ============================
  // 初期化
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();

  // ============================
  // 宅配会社リスト
  // ============================
  function setupCourierOptions() {
    const options = [
      { value: "yamato", text: "ヤマト運輸（B2クラウド）" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa", text: "佐川急便（e飛伝Ⅱ）" },
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
      overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">変換中...</div></div>`;
      document.body.appendChild(overlay);
    }
    overlay.style.display = show ? "flex" : "none";
  }

  // ============================
  // 送り主情報
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
  // クレンジング関数群
  // ============================
  function applyCleaning(value, type) {
    if (!value) return "";
    let cleaned = String(value).trim();

    if (type === "tel" || type === "postal") {
      cleaned = cleaned.replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "");
    }
    if (type === "order") {
      cleaned = cleaned.replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "").trim();
    }
    return cleaned;
  }

  // ============================
  // 住所分割
  // ============================
  function splitAddress(address) {
    if (!address) return { pref: "", city: "", rest: "" };
    const prefs = [
      "北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県",
      "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
      "新潟県","富山県","石川県","福井県","山梨県","長野県",
      "岐阜県","静岡県","愛知県","三重県",
      "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
      "鳥取県","島根県","岡山県","広島県","山口県",
      "徳島県","香川県","愛媛県","高知県",
      "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"
    ];
    const pref = prefs.find(p => address.startsWith(p)) || "";
    const rest = address.replace(pref, "");
    const [city, ...restParts] = rest.split(/(?<=市|区|町|村)/);
    return { pref, city: city || "", rest: restParts.join("") };
  }

  function split25(text) {
    if (!text) return ["", ""];
    return [text.slice(0, 25), text.slice(25, 50)];
  }

  // ============================
  // ヤマト運輸変換処理
  // ============================
  async function mergeToYamatoTemplate(csvFile, templateUrl, sender) {
    const text = await csvFile.text();
    const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
    const dataRows = rows.slice(1);
    const res = await fetch(templateUrl);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets["外部データ取り込み基本レイアウト"];

    let rowExcel = 2;
    for (const r of dataRows) {
      const orderNumber = applyCleaning(r[1], "order");
      const postal = applyCleaning(r[10], "postal");
      const addressFull = r[11] || "";
      const name = r[12] || "";
      const phone = applyCleaning(r[13], "tel");
      const senderAddr = splitAddress(sender.address);

      sheet[`B${rowExcel}`] = { v: "0", t: "s" };
      sheet[`C${rowExcel}`] = { v: "0", t: "s" };
      sheet[`A${rowExcel}`] = { v: orderNumber, t: "s" };
      sheet[`E${rowExcel}`] = { v: new Date().toISOString().slice(0,10).replace(/-/g,"/"), t: "s" };
      sheet[`I${rowExcel}`] = { v: phone, t: "s" };
      sheet[`K${rowExcel}`] = { v: postal, t: "s" };
      sheet[`L${rowExcel}`] = { v: addressFull, t: "s" };
      sheet[`P${rowExcel}`] = { v: name, t: "s" };
      sheet[`Y${rowExcel}`] = { v: sender.name, t: "s" };
      sheet[`T${rowExcel}`] = { v: applyCleaning(sender.phone, "tel"), t: "s" };
      sheet[`V${rowExcel}`] = { v: applyCleaning(sender.postal, "postal"), t: "s" };
      sheet[`W${rowExcel}`] = { v: `${senderAddr.pref}${senderAddr.city}${senderAddr.rest}`, t: "s" };
      sheet[`AB${rowExcel}`] = { v: "ブーケフレーム加工品", t: "s" };
      rowExcel++;
    }

    return wb;
  }

  // ============================
  // 日本郵政 ゆうプリR 変換処理
  // ============================
  async function convertToJapanPost(csvFile, sender) {
    const text = await csvFile.text();
    const rows = text.trim().split(/\r?\n/).map(l => l.split(","));
    const dataRows = rows.slice(1);
    const output = [];

    for (const r of dataRows) {
      const orderNumber = applyCleaning(r[1], "order");
      const postal = applyCleaning(r[10], "postal");
      const addressFull = r[11] || "";
      const name = r[12] || "";
      const phone = applyCleaning(r[13], "tel");
      const addrParts = splitAddress(addressFull);

      const rowOut = [];
      rowOut[7] = name;
      rowOut[10] = postal;
      rowOut[11] = addrParts.pref;
      rowOut[12] = addrParts.city;
      rowOut[13] = addrParts.rest;
      rowOut[15] = phone;
      rowOut[22] = sender.name;
      rowOut[30] = applyCleaning(sender.phone, "tel");
      rowOut[34] = "ブーケフレーム加工品";
      rowOut[49] = orderNumber;
      output.push(rowOut);
    }

    const csvText = output.map(r => r.map(v => `"${v || ""}"`).join(",")).join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv;charset=shift_jis" });
  }

// ============================
// 佐川急便 e飛伝Ⅱ CSV変換処理（列調整版）
// ============================
async function convertToSagawa(csvFile, sender) {
  console.log("🚚 佐川変換処理開始（列位置調整版）");

  const formatRes = await fetch("./formats/sagawaFormat.json");
  const format = await formatRes.json();

  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
  const dataRows = rows.slice(1);

  // 既存フォーマットのヘッダ（全体列数保持）
  const headers = format.columns.map(c => c.header);
  const totalCols = headers.length;
  const output = [];

  for (const row of dataRows) {
    const outRow = new Array(totalCols).fill("");

    // ============================
    // 🧩 基本情報抽出
    // ============================
    const orderNumber = applyCleaning(row[1], "order");   // ご注文番号
    const postal = applyCleaning(row[10], "postal");      // 郵便番号
    const addressFull = row[11] || "";                    // 住所
    const name = row[12] || "";                           // 氏名
    const phone = applyCleaning(row[13], "tel");          // 電話番号

    const senderAddr = splitAddress(sender.address);
    const addrParts = splitAddress(addressFull);

    // ============================
    // 🏠 各列マッピング
    // ============================

    // A列: お届け先コード取得区分
    outRow[0] = "0";

    // C列: お届け先電話番号
    outRow[2] = phone;

    // D列: お届け先郵便番号
    outRow[3] = postal;

    // E列: お届け先住所（都道府県＋市区町村＋番地まで）
    outRow[4] = `${addrParts.pref}${addrParts.city}${addrParts.rest}`;

    // H列: お届け先名称（氏名）
    outRow[7] = name;

    // Q列: ご依頼主電話番号（senderPhone）
    outRow[16] = applyCleaning(sender.phone, "tel");

    // R列: ご依頼主郵便番号（senderPostal）
    outRow[17] = applyCleaning(sender.postal, "postal");

    // S列: ご依頼主住所（senderAddress）
    outRow[18] = senderAddr.pref + senderAddr.city + senderAddr.rest;

    // V列: ご依頼主名称（senderName）
    outRow[21] = sender.name;

    // AE列: 品名（固定値）
    outRow[30] = "ブーケフレーム加工品";

    // BH列: ご注文番号（CSV col2）
    outRow[49] = orderNumber;

    // BI列: 出荷日（今日）
    outRow[50] = new Date().toISOString().slice(0, 10).replace(/-/g, "/");

    output.push(outRow);
  }

  // ============================
  // CSV組み立て（SJIS出力・BOMなし）
  // ============================
  const csvText = [headers.join(",")]
    .concat(output.map(r => r.map(v => `"${v || ""}"`).join(",")))
    .join("\r\n");

  const sjisArray = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
  return new Blob([new Uint8Array(sjisArray)], { type: "text/csv;charset=shift_jis" });
}


  // ============================
  // ボタンイベント
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file = fileInput.files[0];
      const courier = courierSelect.value;
      if (!file) return;

      showLoading(true);
      showMessage("変換処理中...", "info");

      try {
        const sender = getSenderInfo();
        if (courier === "japanpost") {
          convertedCSV = await convertToJapanPost(file, sender);
          showMessage("✅ 日本郵政（ゆうプリR）変換完了", "success");
        } else if (courier === "sagawa") {
          convertedCSV = await convertToSagawa(file, sender);
          showMessage("✅ 佐川急便（e飛伝Ⅱ）変換完了", "success");
        } else {
          mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
          showMessage("✅ ヤマト運輸（B2クラウド）変換完了", "success");
        }

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
  // ダウンロード処理
  // ============================
  function setupDownloadButton() {
    downloadBtn.addEventListener("click", () => {
      if (mergedWorkbook) {
        XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
      } else if (convertedCSV) {
        const courier = courierSelect.value;
        const filename =
          courier === "japanpost" ? "yupack_import.csv" :
          courier === "sagawa" ? "sagawa_import.csv" :
          "output.csv";
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
      } else {
        alert("変換データがありません。");
      }
    });
  }
})();
