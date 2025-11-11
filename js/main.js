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
  function cleanTelPostal(v) {
    if (!v) return "";
    return String(v).replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "").trim();
  }

  function cleanOrderNumber(v) {
    if (!v) return "";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "").trim();
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
  // ゆうプリR変換処理
  // ============================
  async function convertToJapanPost(csvFile, sender) {
    const text = await csvFile.text();
    const rows = text.trim().split(/\r?\n/).map(l => l.split(","));
    const dataRows = rows.slice(1);
    const output = [];

    for (const r of dataRows) {
      const orderNumber = cleanOrderNumber(r[1]);
      const postal = cleanTelPostal(r[10]);
      const addressFull = r[11] || "";
      const name = r[12] || "";
      const phone = cleanTelPostal(r[13]);
      const addrParts = splitAddress(addressFull);

      const rowOut = [];
      rowOut[7] = name;
      rowOut[10] = postal;
      rowOut[11] = addrParts.pref;
      rowOut[12] = addrParts.city;
      rowOut[13] = addrParts.rest;
      rowOut[15] = phone;
      rowOut[22] = sender.name;
      rowOut[30] = cleanTelPostal(sender.phone);
      rowOut[34] = "ブーケフレーム加工品";
      rowOut[49] = orderNumber;

      output.push(rowOut);
    }

    const csvText = output.map(r => r.map(v => `"${v || ""}"`).join(",")).join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
  }

 // ============================
// 佐川急便 e飛伝Ⅱ変換処理（ヘッダ付き）
// ============================
async function convertToSagawa(csvFile, sender) {
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(l => l.split(","));
  const dataRows = rows.slice(1);
  const output = [];

  // ✅ ヘッダ定義（佐川の正式フォーマット）
  const header = [
    "お届け先コード取得区分","お届け先コード","お届け先電話番号","お届け先郵便番号","お届け先住所１",
    "お届け先住所２","お届け先住所３","お届け先名称１","お届け先名称２","お客様管理番号","お客様コード",
    "部署ご担当者コード取得区分","部署ご担当者コード","部署ご担当者名称","荷送人電話番号","ご依頼主コード取得区分",
    "ご依頼主コード","ご依頼主電話番号","ご依頼主郵便番号","ご依頼主住所１","ご依頼主住所２",
    "ご依頼主名称１","ご依頼主名称２","荷姿","品名１","品名２","品名３","品名４","品名５",
    "荷札荷姿","荷札品名１","荷札品名２","荷札品名３","荷札品名４","荷札品名５","荷札品名６","荷札品名７","荷札品名８","荷札品名９","荷札品名１０","荷札品名１１",
    "出荷個数","スピード指定","クール便指定","配達日","配達指定時間帯","配達指定時間（時分）","代引金額","消費税","決済種別","保険金額",
    "指定シール１","指定シール２","指定シール３","営業所受取","SRC区分","営業所受取営業所コード","元着区分","メールアドレス","ご不在時連絡先","出荷日",
    "お問い合せ送り状No.","出荷場印字区分","集約解除指定","編集０１","編集０２","編集０３","編集０４","編集０５","編集０６","編集０７","編集０８","編集０９","編集１０"
  ];

  for (const r of dataRows) {
    const orderNumber = cleanOrderNumber(r[1]); // ご注文番号（B列）
    const postal = cleanTelPostal(r[10]);       // 郵便番号（K列）
    const addressFull = r[11] || "";            // 住所（L列）
    const name = r[12] || "";                   // お届け先氏名（M列）
    const phone = cleanTelPostal(r[13]);        // 電話番号（N列）
    const addrParts = splitAddress(addressFull);
    const senderParts = splitAddress(sender.address);

    // ✅ 住所25文字ごとに分割
    const split25 = (txt) => {
      if (!txt) return ["", ""];
      return [txt.slice(0, 25), txt.slice(25, 50)];
    };
    const [rest1, rest2] = split25(addrParts.rest);
    const [sRest1, sRest2] = split25(senderParts.rest);

    // ✅ 日付
    const d = new Date();
    const today = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;

    // ✅ 出力行（配列の列順に合わせて埋める）
    const rowOut = [];
    rowOut[0] = "0"; // お届け先コード取得区分
    rowOut[1] = "";
    rowOut[1] = phone;
    rowOut[2] = postal;
    rowOut[3] = addrParts.pref + addrParts.city;
    rowOut[4] = rest1;
    rowOut[5] = rest2;
    rowOut[6] = name;
    rowOut[7] = "";
    rowOut[8] = orderNumber;
    rowOut[10] = "";
    rowOut[14] = cleanTelPostal(sender.phone);
    rowOut[15] = "0";
    rowOut[17] = cleanTelPostal(sender.phone);
    rowOut[18] = cleanTelPostal(sender.postal);
    rowOut[19] = senderParts.pref + senderParts.city;
    rowOut[20] = sRest1;
    rowOut[21] = sender.name;
    rowOut[25] = "ブーケ加工品";
    rowOut[40] = "1";
    rowOut[58] = today;

    output.push(rowOut);
  }

  // ✅ CSV組み立て（1行目にヘッダを付与）
  const csvText = [header.join(",")]
    .concat(output.map(r => r.map(v => `"${v || ""}"`).join(",")))
    .join("\r\n");

  const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
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
          showMessage("✅ 日本郵政（ゆうプリR）変換完了", "success");
        } else if (courier === "sagawa") {
          convertedCSV = await convertToSagawa(file, sender);
          mergedWorkbook = null;
          showMessage("✅ 佐川急便（e飛伝Ⅱ）変換完了", "success");
        } else {
          mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
          convertedCSV = null;
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
        const filename = courier === "japanpost"
          ? "yupack_import.csv"
          : courier === "sagawa"
          ? "sagawa_import.csv"
          : "output.csv";
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
