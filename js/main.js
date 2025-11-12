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
  // 初期設定
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();

  // 宅配会社リスト
  function setupCourierOptions() {
    const options = [
      { value: "yamato", text: "ヤマト運輸（B2クラウド）" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa", text: "佐川急便（e飛伝Ⅱ）" }
    ];
    courierSelect.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join("");
  }

  // ファイル選択
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

  // メッセージ表示
  function showMessage(text, type = "info") {
    messageBox.style.display = "block";
    messageBox.textContent = text;
    messageBox.className = "message " + type;
  }

  // ローディング表示
  function showLoading(show) {
    let overlay = document.getElementById("loading");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "loading";
      overlay.className = "loading-overlay";
      overlay.innerHTML =
        `<div class="loading-content"><div class="spinner"></div><div class="loading-text">変換中...</div></div>`;
      document.body.appendChild(overlay);
    }
    overlay.style.display = show ? "flex" : "none";
  }

  // 送り主情報取得
  function getSenderInfo() {
    return {
      name: document.getElementById("senderName").value.trim(),
      postal: document.getElementById("senderPostal").value.trim(),
      address: document.getElementById("senderAddress").value.trim(),
      phone: document.getElementById("senderPhone").value.trim(),
    };
  }

  // クレンジング
  function cleanTelPostal(v) {
    if (!v) return "0";
    return String(v).replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "").trim();
  }
  function cleanOrderNumber(v) {
    if (!v) return "0";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "").trim();
  }

  // 住所分割
  function splitAddress(address) {
    if (!address) return { pref: "", city: "", rest: "", building: "" };
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
    const restFull = restParts.join("");
    const [rest1, building] = restFull.split(/[\s　]+/, 2);
    return { pref, city, rest: rest1 || "", building: building || "" };
  }

// ============================
// 佐川急便 e飛伝Ⅱ CSV変換処理（列ずれ修正版）
// ============================
async function convertToSagawa(csvFile, sender) {
  console.log("🚚 佐川変換処理開始（列ずれ補正＋正規列版）");

  // ✅ JSONマッピング読込
  const formatRes = await fetch("./formats/sagawaFormat.json");
  const format = await formatRes.json();

  // ✅ 入力CSV読込
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
  const dataRows = rows.slice(1); // ヘッダ削除

  // ✅ 出力初期化
  const headers = format.columns.map(c => c.header);
  const totalCols = headers.length;
  const output = [];

  for (const row of dataRows) {
    // --- 空欄初期化（列数に完全一致） ---
    const outRow = Array.from({ length: totalCols }, () => "");

    for (let i = 0; i < format.columns.length; i++) {
      const col = format.columns[i];
      let value = "";

      // --- 固定値 ---
      if (col.value !== undefined) {
        value = (col.value === "TODAY")
          ? `${new Date().getFullYear()}/${String(new Date().getMonth() + 1).padStart(2, "0")}/${String(new Date().getDate()).padStart(2, "0")}`
          : col.value;
      }

      // --- CSV参照 ---
      if (col.source?.startsWith("col")) {
        const idx = parseInt(col.source.replace("col", "")) - 1;
        value = row[idx] || "";
      }

      // --- sender参照 ---
      if (col.source?.startsWith("sender")) {
        const key = col.source.replace("sender", "").toLowerCase();
        value = sender[key] || "";
      }

      // --- 特殊マッピング ---
      if (col.header === "お届け先電話番号") value = cleanTelPostal(row[13] || "");
      if (col.header === "お届け先郵便番号") value = cleanTelPostal(row[10] || "");
      if (["お届け先住所１", "お届け先住所２", "お届け先住所３"].includes(col.header)) {
        const addr = splitAddress(row[11] || "");
        if (col.header === "お届け先住所１") value = addr.pref + addr.city;
        if (col.header === "お届け先住所２") value = addr.rest;
        if (col.header === "お届け先住所３") value = addr.building;
      }
      if (col.header === "お届け先名称１") value = row[12] || "";

      // --- クレンジング ---
      if (col.clean) value = cleanTelPostal(value);

      // ✅ 列ずれ防止：明示的な配置
      outRow[i] = value || "";
    }

    // ✅ 最初の列（A列）に固定値"0"をセット
    outRow[0] = "0";
    output.push(outRow);
  }

  // ✅ CSV出力
  const csvText = [headers.join(",")]
    .concat(output.map(r => r.map(v => `"${v || ""}"`).join(",")))
    .join("\r\n");

  const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
}


  // ============================
  // ボタン処理
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file = fileInput.files[0];
      if (!file) return;
      const courier = courierSelect.value;
      showLoading(true);
      try {
        const sender = getSenderInfo();

        if (courier === "sagawa") {
          convertedCSV = await convertToSagawa(file, sender);
          mergedWorkbook = null;
          showMessage("✅ 佐川急便（e飛伝Ⅱ）変換完了", "success");
        } else {
          showMessage("❌ 今は佐川のみ検証対象です", "error");
          return;
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
      if (convertedCSV) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = "sagawa_import.csv";
        link.click();
        URL.revokeObjectURL(link.href);
      } else {
        alert("変換データがありません。");
      }
    });
  }
})();
