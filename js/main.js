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
  // 宅配会社リスト
  // ============================
  const setupCourierOptions = () => {
    const options = [
      { value: "yamato", text: "ヤマト運輸" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa", text: "佐川急便（今後対応予定）" },
    ];
    courierSelect.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join("");
  };

  // ============================
  // ファイル選択
  // ============================
  const setupFileInput = () => {
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
  };

  // ============================
  // メッセージ表示
  // ============================
  const showMessage = (text, type = "info") => {
    messageBox.style.display = "block";
    messageBox.textContent = text;
    messageBox.className = "message " + type;
  };

  // ============================
  // ローディング表示
  // ============================
  const showLoading = show => {
    let overlay = document.getElementById("loading");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "loading";
      overlay.className = "loading-overlay";
      overlay.innerHTML = `<div class="loading-content"><div class="spinner"></div><div class="loading-text">処理中...</div></div>`;
      document.body.appendChild(overlay);
    }
    overlay.style.display = show ? "flex" : "none";
  };

  // ============================
  // 送り主情報
  // ============================
  const getSenderInfo = () => ({
    name: document.getElementById("senderName").value.trim(),
    postal: document.getElementById("senderPostal").value.trim(),
    address: document.getElementById("senderAddress").value.trim(),
    phone: document.getElementById("senderPhone").value.trim(),
  });

  // ============================
  // クレンジング関数
  // ============================
  const cleanTelPostal = v =>
    !v ? "" : String(v).replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "").trim();

  const cleanOrderNumber = v =>
    !v ? "" : String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "").trim();

  const splitAddress = address => {
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
    return { pref, city, rest: restParts.join("") };
  };

  // ============================
  // 日本郵政マッピング
  // ============================
  const japanPostMapping = [
    { col: 1, rule: "固定値 1" },
    { col: 2, rule: "固定値 0" },
    { col: 7, rule: "固定値 1" },
    { col: 8, rule: "CSV M列" },
    { col: 11, rule: "CSV K列" },
    { col: 12, rule: "CSV L列" },
    { col: 13, rule: "住所.市区町村" },
    { col: 14, rule: "住所.番地" },
    { col: 15, rule: "住所.建物" },
    { col: 16, rule: "CSV N列" },
    { col: 17, rule: "senderName" },
    { col: 26, rule: "senderPostal" },
    { col: 27, rule: "senderAddress" },
    { col: 28, rule: "senderPhone" },
  ];

// ============================
// ゆうプリR変換処理（修正版）
// ============================
async function convertToJapanPost(csvFile, sender) {
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(l => l.split(","));
  const dataRows = rows.slice(1); // 1行目削除

  const output = [];

  // Excel列文字 → 数値(A=0, B=1, ..., Z=25, AA=26)
  const colLetterToIndex = letter => {
    return letter
      .split("")
      .reduce((n, c) => n * 26 + (c.charCodeAt(0) - 65 + 1), 0) - 1;
  };

  for (const r of dataRows) {
    const address = r[12] || ""; // L列（住所）
    const parts = splitAddress(address);
    const rowOut = [];

    japanPostMapping.forEach(m => {
      let val = "";
      const rule = m.rule.trim();

      if (rule.startsWith("固定値")) {
        // 固定値
        val = rule.replace("固定値", "").trim();
      } 
      else if (/CSV\s*([A-Z]+)列/.test(rule)) {
        // CSV列参照（A〜Z, AA〜）
        const colLetter = rule.match(/CSV\s*([A-Z]+)列/)[1];
        const colIndex = colLetterToIndex(colLetter);
        val = r[colIndex] || "";
      } 
      else if (rule.startsWith("sender")) {
        // UIからの送り主情報
        const key = rule.replace("sender", "").toLowerCase();
        val = sender[key] || "";
      } 
      else if (rule.startsWith("住所")) {
        // 住所分割
        if (rule.includes("市区町村")) val = parts.city;
        else if (rule.includes("番地")) val = parts.rest;
        else if (rule.includes("建物")) val = "";
      }

      // クレンジング処理（電話・郵便番号など）
      if (m.col === 11 || m.col === 16 || m.col === 26 || m.col === 28) {
        val = cleanTelPostal(val);
      }

      rowOut[m.col - 1] = val;
    });

    output.push(rowOut);
  }

  // CSV生成
  const csvText = output.map(r => r.map(v => `"${v || ""}"`).join(",")).join("\r\n");

  // Shift_JIS変換
  const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
}


  // ============================
  // 変換ボタン
  // ============================
  const setupConvertButton = () => {
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
  };

  const setupDownloadButton = () => {
    downloadBtn.addEventListener("click", () => {
      if (mergedWorkbook) XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
      else if (convertedCSV) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = "yupack_import.csv";
        link.click();
      } else alert("変換データがありません。");
    });
  };

  // ============================
  // 初期化実行
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();
})();
