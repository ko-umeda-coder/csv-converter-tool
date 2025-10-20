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
  let mapping = {};

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
  // メッセージ
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
  // クレンジング関数
  // ============================
  function cleanTelPostal(value) {
    if (!value) return "";
    return String(value)
      .replace(/^="?/, "")
      .replace(/"$/, "")
      .replace(/[^0-9\-]/g, "")
      .trim();
  }

  // ============================
  // 外部マッピング読込
  // ============================
  async function loadMapping() {
    const res = await fetch("./js/ヤマト.xlsx");
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    mapping = {};
    data.forEach((row, i) => {
      if (!row[0] || i === 0) return;
      mapping[row[0]] = { source: row[1] || "", rule: row[2] || "" };
    });

    console.log("✅ マッピング読込完了:", mapping);
  }

  // ============================
  // 値の取得ロジック
  // ============================
  function getValueFromRule(rule, csvRow, sender, headerMap) {
    if (!rule) return "";

    // 固定値
    if (rule.startsWith("固定値")) {
      return rule.replace("固定値", "").trim();
    }

    // 今日の日付
    if (rule === "TODAY") {
      const d = new Date();
      return `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;
    }

    // UI項目
    if (rule.startsWith("sender")) {
      return sender[rule.replace("sender", "").toLowerCase()] || "";
    }

    // CSV列指定（例：CSV M列）
    const csvMatch = rule.match(/CSV\s*([A-Z]+)列/);
    if (csvMatch) {
      const colLetter = csvMatch[1];
      const colIndex = colLetter.charCodeAt(0) - 65; // A→0
      return csvRow[colIndex] || "";
    }

    return rule;
  }

  // ============================
  // ヤマト変換処理
  // ============================
  async function mergeToYamatoTemplate(csvFile, templateUrl, sender) {
    await loadMapping();

    const csvText = await csvFile.text();
    const rows = csvText.trim().split(/\r?\n/).map(line => line.split(","));
    const dataRows = rows.slice(1);

    const res = await fetch(templateUrl);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets["外部データ取り込み基本レイアウト"];

    let rowExcel = 2;
    for (const r of dataRows) {
      for (const [yamatoCol, def] of Object.entries(mapping)) {
        const value = getValueFromRule(def.source || def.rule, r, sender);
        const cellRef = yamatoCol + rowExcel;
        const cleaned = /電話|郵便番号/.test(yamatoCol) ? cleanTelPostal(value) : value;
        sheet[cellRef] = { v: cleaned, t: "s" };
      }
      rowExcel++;
    }

    return wb;
  }

  // ============================
  // ボタン処理
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
      showMessage("ヤマトマッピングに基づき変換中...", "info");

      try {
        const sender = getSenderInfo();
        mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
        showMessage("✅ 変換完了。ダウンロードできます。", "success");
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
      if (!mergedWorkbook) {
        alert("変換データがありません。");
        return;
      }
      XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
    });
  }
})();
