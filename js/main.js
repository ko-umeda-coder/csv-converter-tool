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
      { value: "", text: "選択してください" },
      { value: "yamato", text: "ヤマト運輸" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa", text: "佐川急便（今後対応予定）" },
    ];
    courierSelect.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join("");
    courierSelect.disabled = false;
    courierSelect.value = "";

    courierSelect.addEventListener("change", () => {
      if (courierSelect.value) {
        console.log("📦 選択された宅配会社:", courierSelect.value);
        convertBtn.disabled = fileInput.files.length === 0;
      } else {
        convertBtn.disabled = true;
      }
    });
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

  function cleanOrderNumber(value) {
    if (!value) return "";
    return String(value)
      .replace(/^(FAX|EC)/i, "") // ✅ FAX・EC削除
      .replace(/[★\[\]\s]/g, "") // ✅ 記号削除
      .trim();
  }

  // ============================
  // 住所分割
  // ============================
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
    const rest = pref ? address.replace(pref, "") : address;
    const [city, ...restParts] = rest.split(/(?<=市|区|町|村)/);
    return { pref, city, rest: restParts.join("") };
  }

  // ============================
  // 外部マッピング読込（日本郵政 F列）
  // ============================
  async function loadMappingJapanPost() {
    const res = await fetch("./js/ゆうプリR_外部データ取込基本レイアウト.xlsx");
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    mapping = {};
    data.forEach((row, i) => {
      if (!row[0] || i === 0) return;
      const val = row[5]; // F列参照
      mapping[row[0]] = { source: (val !== undefined && val !== null) ? String(val).trim() : "" };
    });

    console.log("✅ 日本郵政マッピング読込完了:", mapping);
  }

  // ============================
  // 値取得ロジック（安全型処理）
  // ============================
  function getValueFromRule(rule, csvRow, sender) {
    if (rule == null) return "";
    if (typeof rule !== "string") rule = String(rule);
    rule = rule.trim();

    if (rule.startsWith("固定値")) return rule.replace("固定値", "").trim();
    if (/^\d+$/.test(rule)) return rule; // 固定値 0, 1 など
    if (rule === "TODAY") {
      const d = new Date();
      return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, "0")}/${String(d.getDate()).padStart(2, "0")}`;
    }
    if (rule.startsWith("sender")) return sender[rule.replace("sender", "").toLowerCase()] || "";

    const match = rule.match(/CSV\s*([A-Z]+)列/);
    if (match) {
      const idx = match[1].charCodeAt(0) - 65;
      return csvRow[idx] || "";
    }

    return rule;
  }

  // ============================
  // 日本郵政（ゆうプリR）変換処理
  // ============================
  async function mergeToJapanpostTemplate(csvFile, templateUrl, sender) {
    await loadMappingJapanPost();

    const csvText = await csvFile.text();
    const rows = csvText.trim().split(/\r?\n/).map(line => line.split(","));
    const dataRows = rows.slice(1);

    const res = await fetch(templateUrl);
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];

    let rowExcel = 2;
    for (const r of dataRows) {
      for (const [col, def] of Object.entries(mapping)) {
        // ✅ def.source を安全に参照
        const value = getValueFromRule(def.source, r, sender);
        sheet[`${col}${rowExcel}`] = { v: value, t: "s" };
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
      if (!file || !courier) {
        showMessage("宅配会社を選択してください。", "error");
        return;
      }

      showLoading(true);
      showMessage("変換中...", "info");

      try {
        const sender = getSenderInfo();

        if (courier === "yamato") {
          mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
        } else if (courier === "japanpost") {
          mergedWorkbook = await mergeToJapanpostTemplate(file, "./js/ゆうプリR_外部データ取込基本レイアウト.xlsx", sender);
        } else {
          showMessage("現在対応しているのはヤマト運輸・日本郵政のみです。", "error");
          showLoading(false);
          return;
        }

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

      const courier = courierSelect.value;
      const fileName = courier === "japanpost" ? "japanpost_import.xlsx" : "yamato_b2_import.xlsx";
      XLSX.writeFile(mergedWorkbook, fileName);
    });
  }
})();
