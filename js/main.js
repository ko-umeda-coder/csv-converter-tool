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
      { value: "yamato", text: "ヤマト運輸" },
      { value: "japanpost", text: "日本郵政（WEBゆうプリ）" },
      { value: "sagawa", text: "佐川急便（e飛伝3）" },
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
      overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">処理中...</div></div>`;
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
  function cleanTelPostal(v) {
    if (!v) return "";
    return String(v)
      .replace(/^="?/, "")
      .replace(/"$/, "")
      .replace(/[^0-9\-]/g, "")
      .trim();
  }

  function cleanOrderNumber(v) {
    if (!v) return "";
    return String(v)
      .replace(/^(FAX|EC)/i, "")
      .replace(/[★\[\]\s]/g, "")
      .trim();
  }

  function splitAddress(address) {
    if (!address) return { pref: "", city: "", rest: "" };
    const prefList = [
      "北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県",
      "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
      "新潟県","富山県","石川県","福井県","山梨県","長野県",
      "岐阜県","静岡県","愛知県","三重県",
      "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
      "鳥取県","島根県","岡山県","広島県","山口県",
      "徳島県","香川県","愛媛県","高知県",
      "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"
    ];
    const pref = prefList.find(p => address.startsWith(p)) || "";
    const rest = address.replace(pref, "");
    const [city, ...restParts] = rest.split(/(?<=市|区|町|村)/);
    return { pref, city, rest: restParts.join("") };
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
      const orderNumber = cleanOrderNumber(r[1]);
      const postal = cleanTelPostal(r[10]);
      const addressFull = r[11] || "";
      const name = r[12] || "";
      const phone = cleanTelPostal(r[13]);
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
      sheet[`T${rowExcel}`] = { v: cleanTelPostal(sender.phone), t: "s" };
      sheet[`V${rowExcel}`] = { v: cleanTelPostal(sender.postal), t: "s" };
      sheet[`W${rowExcel}`] = { v: `${senderAddr.pref}${senderAddr.city}${senderAddr.rest}`, t: "s" };
      sheet[`AB${rowExcel}`] = { v: "ブーケフレーム加工品", t: "s" };
      rowExcel++;
    }

    return wb;
  }

  // ============================
  // WEBゆうプリ変換処理
  // ============================
  async function convertToJapanPost(csvFile, sender) {
    const text = await csvFile.text();
    const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
    const res = await fetch("./js/ゆうプリR_外部データ取込基本レイアウト.xlsx");
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: 0, c })];
      headers.push(cell ? String(cell.v).trim() : "");
    }
    const dataRows = rows.slice(1);
    const output = [];

    for (const r of dataRows) {
      const orderNumber = cleanOrderNumber(r[1] || "");
      const postal = cleanTelPostal(r[11] || "");
      const addressFull = r[12] || "";
      const name = r[13] || "";
      const phone = cleanTelPostal(r[14] || "");
      const addrParts = splitAddress(addressFull);
      const senderAddr = splitAddress(sender.address);
      const rowOut = new Array(headers.length).fill("");

      rowOut[0] = "1";
      rowOut[1] = "0";
      rowOut[6] = "1";
      rowOut[8] = "様";
      rowOut[64] = "0";
      rowOut[71] = "0";

      rowOut[7] = name;
      rowOut[10] = postal;
      rowOut[11] = addrParts.pref;
      rowOut[12] = addrParts.city;
      if (addrParts.rest.length > 25) {
        rowOut[13] = addrParts.rest.slice(0, 25);
        rowOut[14] = addrParts.rest.slice(25);
      } else {
        rowOut[13] = addrParts.rest;
        rowOut[14] = "";
      }
      rowOut[15] = phone;
      rowOut[22] = sender.name;
      rowOut[25] = cleanTelPostal(sender.postal);

      const senderAddrParts = splitAddress(sender.address);
      rowOut[26] = senderAddrParts.pref;
      rowOut[27] = senderAddrParts.city;
      if (senderAddrParts.rest.length > 25) {
        rowOut[28] = senderAddrParts.rest.slice(0, 25);
        rowOut[29] = senderAddrParts.rest.slice(25);
      } else {
        rowOut[28] = senderAddrParts.rest;
        rowOut[29] = "";
      }
      rowOut[30] = cleanTelPostal(sender.phone);
      rowOut[32] = orderNumber;
      rowOut[34] = "ブーケ加工品";

      output.push(rowOut);
    }

    const csvText = output.map(row => row.map(v => `"${v ?? ""}"`).join(",")).join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
  }

  // ============================
  // 佐川急便（e飛伝3）変換処理
  // ============================
  async function convertToSagawa(csvFile, sender) {
    const text = await csvFile.text();
    const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
    const dataRows = rows.slice(1);
    const res = await fetch("./js/sagawa_template.xlsx");
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];

    let rowExcel = 2;
    for (const r of dataRows) {
      const orderNumber = cleanOrderNumber(r[1]);
      const postal = cleanTelPostal(r[10] || r[11]);
      const addressFull = r[11] || r[12];
      const name = r[12] || r[13];
      const phone = cleanTelPostal(r[13] || r[14]);
      const addrParts = splitAddress(addressFull);
      const senderAddr = splitAddress(sender.address);

      ws[`C${rowExcel}`] = { v: phone, t: "s" };
      ws[`D${rowExcel}`] = { v: postal, t: "s" };
      ws[`E${rowExcel}`] = { v: addrParts.pref, t: "s" };
      ws[`F${rowExcel}`] = { v: addrParts.city, t: "s" };
      ws[`G${rowExcel}`] = { v: addrParts.rest, t: "s" };
      ws[`H${rowExcel}`] = { v: name, t: "s" };
      ws[`J${rowExcel}`] = { v: orderNumber, t: "s" };
      ws[`R${rowExcel}`] = { v: cleanTelPostal(sender.phone), t: "s" };
      ws[`S${rowExcel}`] = { v: cleanTelPostal(sender.postal), t: "s" };
      ws[`T${rowExcel}`] = { v: senderAddr.pref, t: "s" };
      ws[`U${rowExcel}`] = { v: senderAddr.city + senderAddr.rest, t: "s" };
      ws[`V${rowExcel}`] = { v: sender.name, t: "s" };
      ws[`Y${rowExcel}`] = { v: "ブーケ加工品", t: "s" };
      ws[`AQ${rowExcel}`] = { v: 1, t: "n" };
      ws[`BO${rowExcel}`] = { v: new Date().toISOString().slice(0, 10).replace(/-/g, "/"), t: "s" };

      rowExcel++;
    }

    return wb;
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
          mergedWorkbook = null;
          showMessage("✅ ゆうプリR変換完了", "success");
        } else if (courier === "sagawa") {
          mergedWorkbook = await convertToSagawa(file, sender);
          convertedCSV = null;
          showMessage("✅ 佐川急便変換完了", "success");
        } else {
          mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
          convertedCSV = null;
          showMessage("✅ ヤマト変換完了", "success");
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

  function setupDownloadButton() {
    downloadBtn.addEventListener("click", () => {
      if (mergedWorkbook) {
        const courier = courierSelect.value;
        let fileName = "output.xlsx";
        if (courier === "yamato") fileName = "yamato_b2_import.xlsx";
        else if (courier === "sagawa") fileName = "sagawa_ehiden_import.xlsx";
        XLSX.writeFile(mergedWorkbook, fileName);
      } else if (convertedCSV) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = "yupack_import.csv";
        link.click();
        URL.revokeObjectURL(link.href);
      } else {
        alert("変換データがありません。");
      }
    });
  }
})();
