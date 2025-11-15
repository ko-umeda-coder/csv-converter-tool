// ============================
// XLSXライブラリ読み込み待機
// ============================
const waitForXLSX = () => new Promise((resolve) => {
  const check = () => {
    if (window.XLSX) resolve();
    else setTimeout(check, 50);
  };
  check();
});

// ============================
// 都道府県リスト（全国47）
// ============================
const PREFS = [
  "北海道","青森県","岩手県","宮城県","秋田県","山形県","福島県",
  "茨城県","栃木県","群馬県","埼玉県","千葉県","東京都","神奈川県",
  "新潟県","富山県","石川県","福井県","山梨県","長野県",
  "岐阜県","静岡県","愛知県","三重県",
  "滋賀県","京都府","大阪府","兵庫県","奈良県","和歌山県",
  "鳥取県","島根県","岡山県","広島県","山口県",
  "徳島県","香川県","愛媛県","高知県",
  "福岡県","佐賀県","長崎県","熊本県","大分県","宮崎県","鹿児島県","沖縄県"
];

// ============================
// 都道府県 + 市区町村以下を分離
// ============================
function splitAddressPref(addr) {
  if (!addr) return ["", ""];

  // Trim + 全角/半角スペース除去
  const a = addr.trim().replace(/^[ 　]+/, "");

  for (const pref of PREFS) {
    if (a.startsWith(pref)) {
      return [pref, a.slice(pref.length)];
    }
  }
  return ["", a];
}

// ============================
// 文字列を固定長で分割
// ============================
function splitByLength(text, partLen, maxParts) {
  const s = text || "";
  const parts = [];
  for (let i = 0; i < maxParts; i++) {
    const start = i * partLen;
    if (start >= s.length) {
      parts.push("");
    } else {
      parts.push(s.slice(start, start + partLen));
    }
  }
  return parts;
}

// ============================
// CSVを安全に読み込む（XLSXパーサ）
// ============================
function parseCsvSafe(csvText) {
  const wb = XLSX.read(csvText, { type: "string" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1 });
}

// ============================
// メイン処理
// ============================
(async () => {
  await waitForXLSX();
  console.log("🔥 main.js 起動（完全版）");

  const fileInput     = document.getElementById("csvFile");
  const fileWrapper   = document.getElementById("fileWrapper");
  const fileName      = document.getElementById("fileName");
  const convertBtn    = document.getElementById("convertBtn");
  const downloadBtn   = document.getElementById("downloadBtn");
  const messageBox    = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;
  let convertedCSV   = null;

  // ============================
  // 初期化
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();

  function setupCourierOptions() {
    courierSelect.innerHTML = `
      <option value="yamato">ヤマト運輸（B2クラウド）</option>
      <option value="japanpost">日本郵政（ゆうプリR）</option>
      <option value="sagawa">佐川急便（e飛伝Ⅱ）</option>`;
  }

  function getSenderInfo() {
    return {
      name:    document.getElementById("senderName").value.trim(),
      postal:  cleanTelPostal(document.getElementById("senderPostal").value.trim()),
      address: document.getElementById("senderAddress").value.trim(),
      phone:   cleanTelPostal(document.getElementById("senderPhone").value.trim()),
    };
  }

  // ============================
  // UI 周り
  // ============================
  function setupFileInput() {
    fileInput.addEventListener("change", () => {
      if (fileInput.files.length > 0) {
        fileName.textContent = fileInput.files[0].name;
        fileWrapper.classList.add("has-file");
        convertBtn.disabled = false;
      } else {
        fileName.textContent = "";
        fileWrapper.classList.remove("has-file");
        convertBtn.disabled = true;
      }
    });
  }

  function showMessage(text, type = "info") {
    messageBox.style.display = "block";
    messageBox.textContent = text;
    messageBox.className = "message " + type;
  }

  function showLoading(show) {
    let overlay = document.getElementById("loading");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "loading";
      overlay.className = "loading-overlay";
      overlay.innerHTML = `
        <div class="loading-content">
          <div class="spinner"></div>
          <div class="loading-text">変換中...</div>
        </div>`;
      document.body.appendChild(overlay);
    }
    overlay.style.display = show ? "flex" : "none";
  }

  // ============================
  // 共通ユーティリティ
  // ============================
  function cleanTelPostal(v) {
    if (!v) return "";
    return String(v).replace(/[^0-9\-]/g, "");
  }

  function cleanOrderNumber(v) {
    if (!v) return "";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "");
  }

  // ==========================================================
  // 🟥 ゆうパック（都道府県 + 市区町村以下25×3）
  // ==========================================================
  async function convertToJapanPost(csvFile, sender) {
    console.log("📮 ゆうパック変換開始（完全版）");

    const csvText = await csvFile.text();
    const rows = parseCsvSafe(csvText);
    const data = rows.slice(1);

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    const output = [];

    // ご依頼主住所（都道府県＋市区町村以下25×3）
    const [senderPref, senderRest] = splitAddressPref(sender.address);
    const senderRestLines = splitByLength(senderRest, 25, 3);
    const senderAddrLines = [senderPref, ...senderRestLines];

    for (const r of data) {
      const name    = r[14] || "";
      const postal  = cleanTelPostal(r[10] || "");
      const addrRaw = r[11] || "";
      const phone   = cleanTelPostal(r[13] || "");
      const orderNo = cleanOrderNumber(r[1] || "");

      // 住所（都道府県＋市区町村以下25×3）
      const [pref, rest] = splitAddressPref(addrRaw);
      const restLines = splitByLength(rest, 25, 3);
      const toAddrLines = [pref, ...restLines];

      const row = [];

      // 必須列
      row.push("1","0","","","","","1");

      row.push(name, "様", "", postal);

      // 都道府県 + 市区町村以下
      row.push(toAddrLines[0], toAddrLines[1], toAddrLines[2], toAddrLines[3]);

      row.push(phone, "", "", "");
      row.push("","","");

      // ご依頼主
      row.push(sender.name, "", "", sender.postal);
      row.push(senderAddrLines[0], senderAddrLines[1], senderAddrLines[2], senderAddrLines[3]);
      row.push(sender.phone, "", orderNo, "");

      // 品名
      row.push("ブーケ加工品","","");

      // 日付
      row.push(todayStr,"","","","","");

      while (row.length < 64) row.push("");
      row.push("0"); // 65 割引
      while (row.length < 71) row.push("");
      row.push("0"); // 72 完了通知

      output.push(row);
    }

    const csvOut = output
      .map(r => r.map(v => `"${v ?? ""}"`).join(","))
      .join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvOut), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
  }

  // ==========================================================
  // 🟩 佐川（25文字 × 3 分割）※従来仕様
  // ==========================================================
  async function convertToSagawa(csvFile, sender) {
    console.log("📦 佐川変換開始（従来仕様）");

    const csvText = await csvFile.text();
    const rows = parseCsvSafe(csvText);
    const data = rows.slice(1);

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    const output = [];

    const senderAddrLines = splitByLength(sender.address, 25, 2);

    for (const r of data) {
      const out = Array(74).fill("");

      const addrFull = r[12] || "";
      const postal   = cleanTelPostal(r[11] || "");
      const tel      = cleanTelPostal(r[14] || "");
      const name     = r[13] || "";
      const orderNo  = cleanOrderNumber(r[1] || "");

      const toAddrLines = splitByLength(addrFull, 25, 3);

      out[0] = "0";
      out[2] = tel;
      out[3] = postal;

      out[4] = toAddrLines[0];
      out[5] = toAddrLines[1];
      out[6] = toAddrLines[2];

      out[7] = name;
      out[25] = orderNo;

      out[17] = sender.phone;
      out[18] = sender.postal;
      out[19] = senderAddrLines[0];
      out[20] = senderAddrLines[1];
      out[21] = sender.name;

      out[24] = "ブーケ加工品";
      out[58] = todayStr;

      output.push(out);
    }

    const csvTextOut =
      output.map(r => r.map(v => `"${v ?? ""}"`).join(",")).join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }

  // ==========================================================
  // 🟦 ヤマト（25字 × 2）※従来仕様
  // ==========================================================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚 ヤマト変換開始（従来仕様）");

    const csvText = await csvFile.text();
    const rows = parseCsvSafe(csvText);
    const data = rows.slice(1);

    const res = await fetch("./js/newb2web_template1.xlsx");
    const wb = XLSX.read(await res.arrayBuffer(),{type:"array"});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const header = XLSX.utils.sheet_to_json(sheet,{header:1})[0];

    function colLetter(i){
      let s=""; 
      while(i>=0){ s=String.fromCharCode(i%26+65)+s; i=Math.floor(i/26)-1; }
      return s;
    }
    function idx(key){
      return header.findIndex(h=>typeof h==="string"&&h.includes(key));
    }

    const map = {
      order : idx("お客様管理番号"),
      type  : idx("送り状種類"),
      cool  : idx("クール区分"),
      ship1 : idx("出荷予定日"),
      ship2 : idx("出荷日"),
      tel   : idx("お届け先電話番号"),
      zip   : idx("お届け先郵便番号"),
      adr   : idx("お届け先住所"),
      apt   : idx("お届け先アパートマンション"),
      name  : idx("お届け先名"),
      honor : idx("敬称"),
      sTel  : idx("ご依頼主電話番号"),
      sZip  : idx("ご依頼主郵便番号"),
      sAdr  : idx("ご依頼主住所"),
      sApt  : idx("ご依頼主アパートマンション"),
      sName : idx("ご依頼主名"),
      item  : idx("品名１")
    };

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    let rowExcel = 2;

    function set(i,val){
      if(i < 0) return;
      sheet[colLetter(i)+rowExcel] = { v: val, t: "s" };
    }

    for(const r of data){
      const order = cleanOrderNumber(r[1]  || "");
      const tel   = cleanTelPostal(r[14]   || "");
      const zip   = cleanTelPostal(r[11]   || "");
      const adr   = r[12] || "";
      const name  = r[13] || "";

      const toAddrLines = splitByLength(adr, 25, 2);
      const senderAddrLines = splitByLength(sender.address, 25, 2);

      set(map.order, order);
      set(map.type, "0");
      set(map.cool, "0");
      set(map.ship1, todayStr);
      set(map.ship2, todayStr);

      set(map.tel, tel);
      set(map.zip, zip);

      set(map.adr, toAddrLines[0]);
      set(map.apt, toAddrLines[1]);

      set(map.name, name);
      set(map.honor, "様");

      set(map.sTel, sender.phone);
      set(map.sZip, sender.postal);
      set(map.sAdr, senderAddrLines[0]);
      set(map.sApt, senderAddrLines[1]);
      set(map.sName, sender.name);

      set(map.item, "ブーケ加工品");

      rowExcel++;
    }

    return wb;
  }

  // ============================
  // 変換ボタン
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file    = fileInput.files[0];
      const courier = courierSelect.value;
      if (!file) return;

      const sender = getSenderInfo();
      showLoading(true);

      try {
        if (courier === "yamato") {
          mergedWorkbook = await convertToYamato(file, sender);
          convertedCSV   = null;
        } else if (courier === "japanpost") {
          convertedCSV   = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
        } else { // sagawa
          convertedCSV   = await convertToSagawa(file, sender);
          mergedWorkbook = null;
        }

        showMessage("✔ 変換完了（完全版）", "success");
        downloadBtn.style.display = "block";

      } catch (e) {
        console.error(e);
        showMessage("変換中にエラーが発生しました。", "error");
      } finally {
        showLoading(false);
      }
    });
  }

  // ============================
  // ダウンロード
  // ============================
  function setupDownloadButton() {
    downloadBtn.addEventListener("click", () => {
      const courier = courierSelect.value;

      if (courier === "yamato" && mergedWorkbook) {
        XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
        return;
      }

      if (convertedCSV) {
        const name =
          courier === "japanpost" ? "yupack_import.csv" :
          courier === "sagawa"    ? "sagawa_import.csv" :
          "output.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = name;
        link.click();
        URL.revokeObjectURL(link.href);
      }
    });
  }

})();
