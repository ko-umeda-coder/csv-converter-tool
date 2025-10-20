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
// main.js本体を安全に実行
// ============================
(async () => {
  await waitForXLSX();
  console.log("✅ main.js 初期化開始");

  // ============================
  // DOM要素取得
  // ============================
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
  // 初期設定
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();

  function setupCourierOptions() {
    const options = [
      { value: "yamato", text: "ヤマト運輸" },
      { value: "sagawa", text: "佐川急便（今後対応予定）" },
      { value: "japanpost", text: "日本郵政（今後対応予定）" },
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
      overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">処理中です...</div></div>`;
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
  // クレンジング関数群
  // ============================
  function cleanTelPostal(value) {
    if (!value) return "";
    return String(value)
      .replace(/^="?/, "")   // =" の除去
      .replace(/"$/, "")     // 末尾の " の除去
      .replace(/[^0-9\-]/g, ""); // 数字とハイフン以外削除
  }

  function cleanOrderNumber(value) {
    if (!value) return "";
    return String(value)
      .replace(/^FAX/i, "")       // "FAX"除去
      .replace(/[★\[\]\s]/g, "")  // 特殊文字削除
      .trim();
  }

  function splitAddress(address) {
    if (!address) return { pref: "", city: "", rest: "" };
    const prefectures = ["北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県",
      "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
      "新潟県","富山県","石川県","福井県","山梨県","長野県",
      "岐阜県","静岡県","愛知県","三重県",
      "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
      "鳥取県","島根県","岡山県","広島県","山口県",
      "徳島県","香川県","愛媛県","高知県",
      "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"];
    const pref = prefectures.find(p => address.startsWith(p)) || "";
    const rest = pref ? address.replace(pref, "") : address;
    const [city, ...restParts] = rest.split(/(?<=市|区|町|村)/);
    return { pref, city, rest: restParts.join("") };
  }

  // ============================
  // メイン変換処理
  // ============================
  async function mergeToYamatoTemplate(csvFile, templateUrl, sender) {
    const csvText = await csvFile.text();
    const rows = csvText.trim().split(/\r?\n/).map(line => line.split(","));
    const dataRows = rows.slice(1);

    const response = await fetch(templateUrl);
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    const sheet = workbook.Sheets["外部データ取り込み基本レイアウト"];

    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, "0");
    const dd = String(today.getDate()).padStart(2, "0");
    const shipDate = `${yyyy}/${mm}/${dd}`;

    let rowExcel = 2;
    for (const r of dataRows) {
      const orderNumber = cleanOrderNumber(r[1]);
      const postal = cleanTelPostal(r[10]);
      const phone = cleanTelPostal(r[13]);
      const addressFull = r[11] || "";
      const name = r[12] || "";

      const addrParts = splitAddress(addressFull);
      const senderAddrParts = splitAddress(sender.address);

      sheet[`A${rowExcel}`] = { v: orderNumber, t: "s" };
      sheet[`E${rowExcel}`] = { v: shipDate, t: "s" };
      sheet[`K${rowExcel}`] = { v: postal, t: "s" };
      sheet[`L${rowExcel}`] = { v: `${addrParts.pref}${addrParts.city}${addrParts.rest}`, t: "s" };
      sheet[`P${rowExcel}`] = { v: name, t: "s" };
      sheet[`I${rowExcel}`] = { v: phone, t: "s" };
      sheet[`Y${rowExcel}`] = { v: sender.name, t: "s" };
      sheet[`T${rowExcel}`] = { v: cleanTelPostal(sender.phone), t: "s" };
      sheet[`V${rowExcel}`] = { v: cleanTelPostal(sender.postal), t: "s" };
      sheet[`W${rowExcel}`] = { v: `${senderAddrParts.pref}${senderAddrParts.city}${senderAddrParts.rest}`, t: "s" };
      sheet[`AB${rowExcel}`] = { v: "ブーケフレーム加工品", t: "s" };

      rowExcel++;
    }
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
      showMessage("ヤマトテンプレートに転記中...", "info");

      try {
        const sender = getSenderInfo();
        const templatePath = "./js/newb2web_template1.xlsx";
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
})();
