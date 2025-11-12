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
  console.log("🚚 佐川変換処理開始（列ずれ補正＋明示列版）");

  // テンプレートのヘッダー列数を取得（JSONが72列あることを想定）
  // ※JSONの取得に失敗した場合、デフォルトで72列（A～BT）として処理を継続
  let totalCols = 72;
  let headers = [];
  try {
    const formatRes = await fetch("./formats/sagawaFormat.json");
    const format = await formatRes.json();
    totalCols = format.columns ? format.columns.length : 72;
    headers = format.columns ? format.columns.map(c => c.header) : [];
  } catch (e) {
    console.error("formats/sagawaFormat.jsonの読み込みに失敗しました。", e);
    // ヘッダーは空のまま処理を続行（CSV出力時にデータのみになるが列位置は担保）
  }

  // 入力CSV読込
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
  const dataRows = rows.slice(1); // ヘッダ削除

  const output = [];

  // 送り主住所を結合 (正しい版のT, U列に格納するため)
  const senderAddr = splitAddress(sender.address);
  const senderAddressCombined = senderAddr.pref + senderAddr.city + senderAddr.rest + senderAddr.building;

  for (const row of dataRows) {
    // --- 空欄初期化（列数に完全一致） ---
    const outRow = Array.from({ length: totalCols }, () => "");

    // ============================
    // 🧩 入力CSVからのデータ抽出とクレンジング
    // ============================
    const orderNumber = cleanOrderNumber(row[1] || "");   // ご注文番号 (入力CSV col 2)
    const name = row[12] || "";                           // 氏名 (入力CSV col 13)
    const phone = cleanTelPostal(row[13] || "");          // 電話番号 (入力CSV col 14)
    const postal = cleanTelPostal(row[10] || "");         // 郵便番号 (入力CSV col 11)
    const addressFull = row[11] || "";                    // 住所 (入力CSV col 12)

    // 住所分割
    const addrParts = splitAddress(addressFull);

    // ============================
    // 🏠 明示的な列マッピング (正しい版に合わせたインデックス)
    // ============================

    // A列 (0): お届け先コード取得区分
    outRow[0] = "0"; // 必須

    // C列 (2): お届け先電話番号
    outRow[2] = phone;

    // D列 (3): お届け先郵便番号
    outRow[3] = postal;

    // E列 (4): お届け先住所１ (都道府県＋市区町村)
    outRow[4] = addrParts.pref + addrParts.city;

    // F列 (5): お届け先住所２ (番地)
    outRow[5] = addrParts.rest;
    
    // G列 (6): お届け先住所３ (ビル名など)
    outRow[6] = addrParts.building;

    // H列 (7): お届け先名称１（氏名）
    outRow[7] = name;

    // ✅ I列 (8): お届け先名称２（正しい版に合わせ、ここに注文番号を格納）
    outRow[8] = orderNumber;
    
    // -----------------------------------
    // ご依頼主情報
    // -----------------------------------
    
    // R列 (17): ご依頼主電話番号
    outRow[17] = cleanTelPostal(sender.phone);

    // S列 (18): ご依頼主郵便番号
    outRow[18] = cleanTelPostal(sender.postal);

    // T列 (19): ご依頼主住所１
    // U列 (20): ご依頼主住所２
    // 「正しい版」に合わせ、ご依頼主住所は分割せずフルアドレスを格納
    outRow[19] = senderAddressCombined;
    outRow[20] = senderAddressCombined;

    // V列 (21): ご依頼主名称１
    outRow[21] = sender.name;

    // -----------------------------------
    // 品名・日付
    // -----------------------------------

    // AE列 (30): 荷札品名１（固定値）
    outRow[30] = "ブーケフレーム加工品";
    
    // BG列 (58): 出荷日 (YYYY/MM/DD 形式)
    const today = new Date();
    const dateStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, "0")}/${String(today.getDate()).padStart(2, "0")}`;
    outRow[58] = dateStr;

    output.push(outRow);
  }

  // CSV組み立て（SJIS出力・BOMなし）
  const csvText = [headers.join(",")]
    .concat(output.map(r => r.map(v => `"${v || ""}"`).join(",")))
    .join("\r\n");

  // Encodingライブラリの利用 (元のコードに従う)
  const sjisArray = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
  return new Blob([new Uint8Array(sjisArray)], { type: "text/csv" });
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
