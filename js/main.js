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
// 佐川急便 e飛伝Ⅱ CSV変換処理（列ずれ完全修正版）
// ============================
async function convertToSagawa(csvFile, sender) {
  console.log("🚚 佐川変換処理開始（列ずれ完全修正版）");

  // 入力CSV読込
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(line => line.split(","));
  const dataRows = rows.slice(1); // ヘッダ削除

  const totalCols = 72; // 常に72列固定（A〜BV）
  const output = [];

  // 固定ヘッダー（e飛伝Ⅱフォーマット）
  const headers = [
    "お届け先コード取得区分","お届け先コード","お届け先電話番号","お届け先郵便番号","お届け先住所1","お届け先住所2",
    "お届け先住所3","お届け先名称1","お届け先名称2","お客様管理番号","お客様コード","部署ご担当者コード取得区分","部署ご担当者コード",
    "部署ご担当者名称","荷送人電話番号","ご依頼主コード取得区分","ご依頼主コード","ご依頼主電話番号","ご依頼主郵便番号",
    "ご依頼主住所1","ご依頼主住所2","ご依頼主名称1","ご依頼主名称2","荷姿","品名1","品名2","品名3","品名4","品名5",
    "荷札荷姿","荷札品名1","荷札品名2","荷札品名3","荷札品名4","荷札品名5","荷札品名6","荷札品名7","荷札品名8","荷札品名9",
    "荷札品名10","荷札品名11","出荷個数","スピード指定","クール便指定","配達日","配達指定時間帯","配達指定時間（時分）","代引金額",
    "消費税","決済種別","保険金額","指定シール1","指定シール2","指定シール3","営業所受取","SRC区分","営業所受取営業所コード",
    "元着区分","メールアドレス","ご不在時連絡先","出荷日","お問い合せ送り状No.","出荷場印字区分","集約解除指定",
    "編集01","編集02","編集03","編集04","編集05","編集06","編集07","編集08","編集09","編集10"
  ];

  // 送り主住所を結合
  const senderAddr = splitAddress(sender.address);
  const senderAddressCombined =
    senderAddr.pref + senderAddr.city + senderAddr.rest + senderAddr.building;

  for (const row of dataRows) {
    const outRow = Array.from({ length: totalCols }, () => "");

    // 入力CSV参照
    const orderNumber = cleanOrderNumber(row[1] || "");
    const name = row[12] || "";
    const phone = cleanTelPostal(row[13] || "");
    const postal = cleanTelPostal(row[10] || "");
    const addressFull = row[11] || "";
    const addrParts = splitAddress(addressFull);

    // 明示マッピング
    outRow[0] = "0"; // お届け先コード取得区分
    outRow[2] = phone; // お届け先電話番号
    outRow[3] = postal; // 郵便番号
    outRow[4] = addrParts.pref + addrParts.city; // 住所1
    outRow[5] = addrParts.rest; // 住所2
    outRow[6] = addrParts.building; // 住所3
    outRow[7] = name; // お届け先名称1
    outRow[8] = orderNumber; // 名称2に注文番号

    // ご依頼主情報
    outRow[17] = cleanTelPostal(sender.phone);
    outRow[18] = cleanTelPostal(sender.postal);
    outRow[19] = senderAddressCombined;
    outRow[20] = senderAddressCombined;
    outRow[21] = sender.name;

    // 品名・出荷日
    outRow[30] = "ブーケフレーム加工品";
    const today = new Date();
    outRow[58] = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, "0")}/${String(today.getDate()).padStart(2, "0")}`;

    output.push(outRow);
  }

  // ✅ ヘッダー＋72列固定のCSV出力
  const csvText = [headers.join(",")]
    .concat(output.map(r => r.map(v => `"${v}"`).join(",")))
    .join("\r\n");

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
