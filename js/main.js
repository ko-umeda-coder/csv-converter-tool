// ============================
// XLSXライブラリ読み込み待機
// ============================
const waitForXLSX = () => new Promise((resolve) => {
  const check = () => {
    if (window.XLSX) {
      console.log("✅ XLSX 読み込み完了");
      resolve();
    } else {
      setTimeout(check, 50);
    }
  };
  check();
});

// ============================
// メイン処理
// ============================
(async () => {
  await waitForXLSX();
  console.log("✅ main.js 起動");

  const fileInput     = document.getElementById("csvFile");
  const fileWrapper   = document.getElementById("fileWrapper");
  const fileName      = document.getElementById("fileName");
  const convertBtn    = document.getElementById("convertBtn");
  const downloadBtn   = document.getElementById("downloadBtn");
  const messageBox    = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;   // ヤマト用（Excel）
  let convertedCSV   = null;   // ゆうパック/佐川用（CSV Blob）

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
      { value: "yamato",    text: "ヤマト運輸（B2クラウド）" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa",    text: "佐川急便（e飛伝Ⅱ）" },
    ];
    courierSelect.innerHTML = options
      .map(o => `<option value="${o.value}">${o.text}</option>`)
      .join("");
  }

  // ============================
  // 送り主情報
  // ============================
  function getSenderInfo() {
    return {
      name:    document.getElementById("senderName").value.trim(),
      postal:  cleanTelPostal(document.getElementById("senderPostal").value.trim()),
      address: document.getElementById("senderAddress").value.trim(),
      phone:   cleanTelPostal(document.getElementById("senderPhone").value.trim()),
    };
  }

  // ============================
  // ファイル入力
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
  // ローディング
  // ============================
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
  // 共通クレンジング
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

// =======================================================
// 住所分割：3社共通 → 都道府県 / 市区町村 / 丁番地＋建物（25文字分割）
// =======================================================
function splitAddress2(address) {
  if (!address) {
    return {
      pref: "",        // 都道府県
      city: "",        // 市区町村
      addr2: "",       // 丁目番地＋建物 25文字以内
      addr3: ""        // addr2 の続き
    };
  }

  // 都道府県一覧
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

  // 都道府県
  const pref = prefs.find(p => address.startsWith(p)) || "";
  let rest = pref ? address.slice(pref.length) : address;

  // 市区町村で分割（市/区/町/村 の直後で分割）
  const cityMatch = rest.match(/^(.*?[市区町村])/);
  const city = cityMatch ? cityMatch[1] : "";
  rest = city ? rest.slice(city.length) : rest;

  // 残り = 丁番地 + 建物名（全てまとめる）
  const restFull = rest.trim();

  // ★ 25文字で分割 ★
  let addr2 = "";
  let addr3 = "";

  if (restFull.length <= 25) {
    addr2 = restFull;
    addr3 = "";
  } else {
    addr2 = restFull.slice(0, 25);
    addr3 = restFull.slice(25);
  }

  // 最終的な返り値
  return {
    pref,
    city,
    addr2,
    addr3
  };
}


// =======================================================
// 住所分割：ヤマトB2用（都道府県 / 市区町村 / 残り / 建物名）
// =======================================================
// ※ メイン関数内の splitAddress2 はそのまま残し、ヤマトB2用はここで新しく定義する
function splitAddressYamato(address) {
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
  let rest = pref ? address.slice(pref.length) : address;

  // 市区町村で分割（市/区/町/村 の直後で分割）
  const cityMatch = rest.match(/^(.*?[市区町村])/);
  const city = cityMatch ? cityMatch[1] : "";
  rest = cityMatch ? rest.slice(city.length) : rest;

  // 建物名・号室の抽出 (単純化)
  let building = "";
  const lastCommaIndex = rest.lastIndexOf("号室");
  if (lastCommaIndex !== -1) {
    building = rest.slice(lastCommaIndex - 4).trim(); // 例: 号室の前に建物名の一部を抽出
    // よりシンプルに、建物名とそれ以外を分ける。
    // B2クラウドでは、住所に「市区町村＋番地」まで、アパートマンションに「建物名・号室」を期待することが多いため、
    // ここでは、建物名と判断できるものを末尾から分離するロジックを簡略化し、「残り」をすべて住所に入れることにします。
    
    // B2クラウドの住所は、都道府県、市区郡町村、番地の3つの列に分かれているわけではないため、
    // 実際は「お届け先住所」に「都道府県＋市区町村＋番地」をセットし、
    // 「お届け先アパートマンション」に「建物名・号室」をセットするのが最も安全です。
    
    // 建物名の自動抽出は難しいため、ここでは**「お届け先住所」に都道府県から番地まで、「お届け先アパートマンション」に建物名・号室をセット**する最も一般的な手法を採用します。
    
    // ただし、元のコードにある `splitAddress2`の定義が不明確なため、
    // **「お届け先住所」にフルアドレスを、「お届け先アパートマンション」を空欄**とする「正解ファイル」のパターンに合わせるのが最優先です。
    return { 
      fullAddress: address.trim(),
      apartment: "" 
    };
  }

  return { 
    fullAddress: address.trim(),
    apartment: "" 
  };
}


// ============================
// ヤマト B2クラウド（正解ファイル準拠修正版）
// ============================
async function convertToYamato(csvFile, sender) {
  console.log("🚚 ヤマトB2変換開始");

  // 入力CSV読み込み
  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1);

  // テンプレート（正解Excelと同じ構成）
  const res = await fetch("./js/newb2web_template1.xlsx");
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type: "array" });

  const sheetName = wb.SheetNames[0];
  const sheet     = wb.Sheets[sheetName];

  // 1行目ヘッダ取得
  const headerRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headerRow  = headerRows[0] || [];

  // ヘッダ検索（完全一致ではなく「含む」）
  function findHeaderIndex(keyword) {
    return headerRow.findIndex(h => typeof h === "string" && h.includes(keyword));
  }

  // 列番号→A/B/C変換
  function colLetter(idx) {
    let s = "";
    let n = idx;
    while (n >= 0) {
      s = String.fromCharCode((n % 26) + 65) + s;
      n = Math.floor(n / 26) - 1;
    }
    return s;
  }

  // -------------------------------
  // 住所処理（正解ファイル準拠: 住所はフルアドレス、アパートは空欄）
  // -------------------------------
  // ※ 正解ファイルはご依頼主住所が分割されていないため、このパターンに合わせる
  const getAddressParts = (addr) => {
    if (!addr) return { full: "", apt: "" };
    
    // ここでは、全てを「full」に格納し、「apt」は空欄とする（正解ファイル準拠）
    return { full: addr.trim(), apt: "" };
  };

  const senderA = getAddressParts(sender.address);

  // マッピングルール
  const ruleDefs = [
    // お客様管理番号 = CSV B列
    { key: "お客様管理番号", type: "csv", col: 1, clean: "order" },

    // 固定値
    { key: "送り状種類", type: "value", value: "0" },
    { key: "クール区分", type: "value", value: "0" },

    // 日付
    { key: "出荷予定日", type: "today" },
    { key: "出荷日",     type: "today" },

    // お届け先（CSV側：K=10, L=11, M=12, N=13）
    { key: "お届け先電話番号", type: "csv", col: 13, clean: "tel" },
    { key: "お届け先郵便番号", type: "csv", col: 10, clean: "postal" },

    // ⭐修正1: お届け先住所はフルアドレスを格納
    { key: "お届け先住所",   type: "addrFull" }, 

    // ⭐修正2: お届け先アパートマンションは空欄を格納
    { key: "お届け先アパートマンション", type: "addrApt" }, 

    { key: "お届け先名", type: "csv", col: 12 },
    { key: "敬称",      type: "value", value: "様" },

    // ご依頼主（UI入力）
    { key: "ご依頼主電話番号",    type: "senderPhone" },
    { key: "ご依頼主郵便番号",    type: "senderPostal" },

    // ⭐修正3: ご依頼主住所はフルアドレスを格納
    { key: "ご依頼主住所",        type: "senderAddressFull" }, 

    // ⭐修正4: ご依頼主アパートマンションは空欄を格納
    { key: "ご依頼主アパートマンション", type: "senderApt" }, 

    { key: "ご依頼主名",          type: "senderName" },

    // 品名
    { key: "品名１", type: "value", value: "ブーケ加工品" }
  ];

  // ヘッダごとの列番号キャッシュ
  const headIndex = {};
  for (const r of ruleDefs) {
    const idx = findHeaderIndex(r.key);
    if (idx >= 0) headIndex[r.key] = idx;
  }

  const today = new Date();
  const todayStr =
    `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  let excelRow = 2; // 2行目から

  // -------------------------------
  // 書き込みループ
  // -------------------------------
  for (const r of data) {

    // お届け先住所の処理（フルアドレスを抽出）
    const addressCsv = r[11] || "";
    const addr = getAddressParts(addressCsv);

    for (const rule of ruleDefs) {
      const idx = headIndex[rule.key];
      if (idx === undefined) continue;

      const col = colLetter(idx);
      const cell = col + excelRow;

      let v = "";

      switch (rule.type) {
        case "value":
          v = rule.value;
          break;

        case "today":
          v = todayStr;
          break;

        case "csv": {
          let src = r[rule.col] || "";
          if (rule.clean === "tel" || rule.clean === "postal")
            src = cleanTelPostal(src);
          if (rule.clean === "order")
            src = cleanOrderNumber(src);
          v = src;
          break;
        }

        // -------------------------------
        // お届け先住所 (修正箇所)
        // -------------------------------
        case "addrFull": // お届け先住所
          v = addr.full;
          break;

        case "addrApt": // お届け先アパートマンション
          v = addr.apt;
          break;

        // -------------------------------
        // 送り主 (修正箇所)
        // -------------------------------
        case "senderPhone":
          v = cleanTelPostal(sender.phone);
          break;

        case "senderPostal":
          v = cleanTelPostal(sender.postal);
          break;

        case "senderAddressFull": // ご依頼主住所
          v = senderA.full;
          break;

        case "senderApt": // ご依頼主アパートマンション
          v = senderA.apt;
          break;

        case "senderName":
          v = sender.name;
          break;
      }

      sheet[cell] = { v, t: "s" };
    }

    excelRow++;
  }

  return wb;
}


// ============================
// ゆうパック（ゆうプリR） 72列固定・ヘッダなし
// ============================
async function convertToJapanPost(csvFile, sender) {
  console.log("📮 ゆうパック（ゆうプリR）変換開始");

  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1); // ヘッダ除去

  const output  = [];

  const today = new Date();
  const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  // ◆ 送り主住所を分割（25文字制限対応）
  const sendAddr = splitAddress2(sender.address);

  for (const r of data) {
    const name        = r[12] || "";                  // M列：氏名
    const postal      = cleanTelPostal(r[10] || "");  // K列：郵便番号
    const addressFull = r[11] || "";                  // L列：住所
    const phone       = cleanTelPostal(r[13] || "");  // N列：電話番号
    const orderNo     = cleanOrderNumber(r[1] || ""); // B列：注文番号

    // ◆ お届け先住所を分割（25文字制限対応）
    const addr = splitAddress2(addressFull);

    const row = [];

    // 👉 ここから 72 列固定で push
    row.push("1");              // 1 商品
    row.push("0");              // 2 着払/代引
    row.push("");               // 3
    row.push("");               // 4
    row.push("");               // 5
    row.push("");               // 6
    row.push("1");              // 7 作成数

    // ★ お届け先
    row.push(name);             // 8 お名前
    row.push("様");             // 9 敬称
    row.push("");               // 10 カナ
    row.push(postal);           // 11 郵便番号
    row.push(addr.pref);        // 12 都道府県
    row.push(addr.city);        // 13 市区町村郡
    row.push(addr.addr2);       // 14 丁番地（25文字制限）
    row.push(addr.addr3);       // 15 建物名（25文字以降）
    row.push(phone);            // 16 電話番号
    row.push("");               // 17 法人名
    row.push("");               // 18 部署
    row.push("");               // 19 メール

    // 20〜22
    row.push("");
    row.push("");
    row.push("");

    // ★ 送り主
    row.push(sender.name);      // 23 ご依頼主名
    row.push("");               // 24 敬称
    row.push("");               // 25 カナ
    row.push(sender.postal);    // 26 郵便番号
    row.push(sendAddr.pref);    // 27 都道府県
    row.push(sendAddr.city);    // 28 市区町村郡
    row.push(sendAddr.addr2);   // 29 丁番地
    row.push(sendAddr.addr3);   // 30 建物名
    row.push(sender.phone);     // 31 電話番号

    row.push("");               // 32 法人名
    row.push(orderNo);          // 33 部署名（注文番号）
    row.push("");               // 34 メール

    row.push("ブーケ加工品");   // 35 品名
    row.push("");               // 36 品名番号
    row.push("");               // 37 個数

    // ★ 発送予定日
    row.push(todayStr);        // 38 発送予定日

    // 39〜72
    for (let i = 39; i <= 72; i++) {
      if (i === 65) row.push("0");     // 割引
      else if (i === 72) row.push("0"); // 配達完了通知（依頼主）
      else row.push("");
    }

    output.push(row);
  }

  // 👉 ヘッダなし・72列の CSV 出力
  const csvOut = output.map(row => row.map(v => `"${v}"`).join(",")).join("\r\n");
  const sjis = Encoding.convert(Encoding.stringToCode(csvOut), "SJIS");

  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
}

  
// ============================
// 佐川 e飛伝Ⅱ（74列固定・住所25文字分割対応）
// ============================
async function convertToSagawa(csvFile, sender) {
  console.log("📦 佐川（e飛伝Ⅱ）変換開始");

  const headers = [
    "お届け先コード取得区分","お届け先コード","お届け先電話番号","お届け先郵便番号",
    "お届け先住所１","お届け先住所２","お届け先住所３",
    "お届け先名称１","お届け先名称２",
    "お客様管理番号","お客様コード","部署ご担当者コード取得区分",
    "部署ご担当者コード","部署ご担当者名称","荷送人電話番号",
    "ご依頼主コード取得区分","ご依頼主コード","ご依頼主電話番号",
    "ご依頼主郵便番号","ご依頼主住所１","ご依頼主住所２",
    "ご依頼主名称１","ご依頼主名称２",
    "荷姿","品名１","品名２","品名３","品名４","品名５",
    "荷札荷姿","荷札品名１","荷札品名２","荷札品名３","荷札品名４","荷札品名５",
    "荷札品名６","荷札品名７","荷札品名８","荷札品名９","荷札品名１０","荷札品名１１",
    "出荷個数","スピード指定","クール便指定","配達日",
    "配達指定時間帯","配達指定時間（時分）","代引金額","消費税","決済種別","保険金額",
    "指定シール１","指定シール２","指定シール３",
    "営業所受取","SRC区分","営業所受取営業所コード","元着区分",
    "メールアドレス","ご不在時連絡先",
    "出荷日","お問い合せ送り状No.","出荷場印字区分","集約解除指定",
    "編集01","編集02","編集03","編集04","編集05",
    "編集06","編集07","編集08","編集09","編集10"
  ];

  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1);

  const today = new Date();
  const todayStr =
    `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  const output = [];

  // ◆ ご依頼主住所分割（25文字制限：addr2 / addr3）
  const sendA = splitAddress2(sender.address);
  const sendAddr1 = (sendA.pref || "") + (sendA.city || "");   // 都道府県 + 市区町村
  const sendAddr2 = (sendA.addr2 || "") + (sendA.addr3 || ""); // 丁番地号 + 建物名（全部）

  for (const r of data) {
    const out = Array(headers.length).fill("");

    const orderNumber = cleanOrderNumber(r[1] || "");
    const postal      = cleanTelPostal(r[10] || "");
    const addressFull = r[11] || "";
    const name        = r[12] || "";
    const phone       = cleanTelPostal(r[13] || "");

    // ★ お届け先住所（25文字制限に分割）
    const addr = splitAddress2(addressFull);

    // ======== ★ 各列へのセット（正解仕様） ========
    out[0]  = "0";                       // A: コード取得区分
    out[2]  = phone;                     // C: 電話番号
    out[3]  = postal;                    // D: 郵便番号
    out[4]  = addr.pref + addr.city;     // E: 住所1（都道府県＋市区町村）
    out[5]  = addr.addr2;                // F: 住所2（25文字以内）
    out[6]  = addr.addr3;                // G: 住所3（残り全部）
    out[7]  = name;                      // H: 名称１（氏名）
    out[8]  = orderNumber;               // I: 名称２（注文番号）

    out[14] = sender.phone;              // O: 荷送人電話番号
    out[17] = sender.phone;              // R: ご依頼主電話番号
    out[18] = sender.postal;             // S: 郵便番号（依頼主）

    // ⭐修正済：住所1 / 住所2 に分割してセット
    out[19] = sendAddr1;                 // T: ご依頼主住所１（都道府県＋市区町村）
    out[20] = sendAddr2;                 // U: ご依頼主住所２（丁目番地号＋建物名）

    out[21] = sender.name;               // V: ご依頼主名称１
    out[25] = "ブーケ加工品";           // Z: 品名１

    out[58] = todayStr;                  // BG: 出荷日（正解どおり）

    output.push(out);
  }

  // CSV書き出し（ヘッダ入り）
  const csvTextOut = [
    headers.join(","),
    ...output.map(row => row.map(v => `"${v}"`).join(","))
  ].join("\r\n");

  const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut), "SJIS");
  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
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
      showMessage("変換処理中...", "info");

      try {
        if (courier === "yamato") {
          mergedWorkbook = await convertToYamato(file, sender);
          convertedCSV   = null;
          showMessage("✅ ヤマトB2用データが完成しました", "success");
        } else if (courier === "japanpost") {
          convertedCSV   = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
          showMessage("✅ ゆうプリR（ゆうパック）用CSVが完成しました", "success");
        } else if (courier === "sagawa") {
          convertedCSV   = await convertToSagawa(file, sender);
          mergedWorkbook = null;
          showMessage("✅ 佐川 e飛伝Ⅱ用CSVが完成しました", "success");
        } else {
          showMessage("未対応の宅配会社です。", "error");
          return;
        }

        downloadBtn.style.display = "block";
        downloadBtn.disabled = false;
      } catch (e) {
        console.error(e);
        showMessage("変換中にエラーが発生しました。", "error");
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
      const courier = courierSelect.value;

      if (courier === "yamato" && mergedWorkbook) {
        XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
        return;
      }

      if (convertedCSV) {
        const filename =
          courier === "japanpost" ? "yupack_import.csv" :
          courier === "sagawa"    ? "sagawa_import.csv" :
          "output.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = filename;
        link.click();
        URL.revokeObjectURL(link.href);
      } else {
        alert("ダウンロード可能なデータがありません。");
      }
    });
  }
})();
