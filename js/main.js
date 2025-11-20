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
// 都道府県を抽出
// ============================
function splitAddressPref(addr) {
  if (!addr) return ["", ""];
  const a = addr.trim().replace(/^[ 　]+/, "");
  for (const pref of PREFS) {
    if (a.startsWith(pref)) return [pref, a.slice(pref.length)];
  }
  return ["", a];
}

// ============================
// 市区町村を抽出
// ============================
function splitCity(addr) {
  if (!addr) return ["", ""];
  const a = addr.trim();
  const match = a.match(/^(.*?[市区町村])/);
  if (match) {
    const city = match[1];
    return [city, a.slice(city.length)];
  }
  return ["", a];
}

// ============================
// 固定長で分割
// ============================
function splitByLength(text, partLen, maxParts) {
  const s = text || "";
  const parts = [];
  for (let i = 0; i < maxParts; i++) {
    const start = i * partLen;
    parts.push(s.slice(start, start + partLen) || "");
  }
  return parts;
}

// ============================
// CSV安全読み込み（UTF-8版 修正版）
// ============================
function parseCsvSafe(csvText) {
  // ここで csvText は「UTF-8 をブラウザが JS 文字列にしたもの」
  // なので、そのまま string として XLSX に渡せばよい
  const wb = XLSX.read(csvText, { type: "string" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false });
}


// ============================
// メイン処理
// ============================
(async () => {
  await waitForXLSX();
  console.log("🔥 main.js 起動（完全版）");

  const fileInput     = document.getElementById("csvFile");
  // 1. 【修正】fileWrapper は index.html に存在しないため、削除
  // const fileWrapper   = document.getElementById("fileWrapper");
  const fileName      = document.getElementById("fileName");
  const convertBtn    = document.getElementById("convertBtn");
  const downloadBtn   = document.getElementById("downloadBtn");
  const messageBox    = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;
  let convertedCSV   = null;

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
      name:    document.getElementById("senderName").value.trim(),
      postal:  cleanTelPostal(document.getElementById("senderPostal").value.trim()),
      address: document.getElementById("senderAddress").value.trim(),
      phone:   cleanTelPostal(document.getElementById("senderPhone").value.trim()),
    };
  }

  // ============================
  // UI
  // ============================
  function setupFileInput() {
    fileInput.addEventListener("change", () => {
      if (fileInput.files.length > 0) {
        fileName.textContent = fileInput.files[0].name;
        // 1. 【修正】fileWrapper に関する処理を削除
        // fileWrapper.classList.add("has-file");
        convertBtn.disabled = false;
      } else {
        fileName.textContent = "";
        // 1. 【修正】fileWrapper に関する処理を削除
        // fileWrapper.classList.remove("has-file");
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

  // 数字とハイフン以外を除去
  let s = String(v).replace(/[^0-9\-]/g, "");

  // ハイフン除去して桁数判定（ゆうプリはハイフンなしで扱う）
  const digits = s.replace(/-/g, "");

  // 13桁を超えたら14桁以降を削除
  if (digits.length > 13) {
    s = digits.slice(0, 13);
  } else {
    s = digits;
  }

  return s;
}


  function cleanOrderNumber(v) {
    return v ? String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "") : "";
  }

// ==========================================================
// 🟥 ゆうパック（Shift-JIS）ゆうプリWEB対応・完全版
// 2. 【修正】文字化け対策のため、BOM付きUTF-8に変更
// ==========================================================


// ----------------------------------------------------------
// ① SJIS 非対応文字の正規化（外字 → 通常字）
// ----------------------------------------------------------
function normalizeForSJIS(str) {
  if (!str) return "";

  let s = String(str);

  const map = {
    "髙": "高", "﨑": "崎", "神": "神", "塚": "塚", "𠮷": "吉",

    "①": "1", "②": "2", "③": "3", "④": "4", "⑤": "5",
    "⑥": "6", "⑦": "7", "⑧": "8", "⑨": "9", "⑩": "10",

    "Ⅰ": "I", "Ⅱ": "II", "Ⅲ": "III",

    "㈱": "(株)", "㈲": "(有)",

    "㎜": "mm", "㎝": "cm", "㎞": "km",
    "㌔": "キロ", "㌢": "センチ", "㌘": "グラム",

    "—": "ー", "–": "ー", "−": "-",

    "’": "'", "”": "\"", "“": "\"",
  };

  for (const [from, to] of Object.entries(map)) {
    s = s.replace(new RegExp(from, "g"), to);
  }

  // サロゲートペア（絵文字等）をすべて削除
  s = s.replace(/[\uD800-\uDFFF]/g, "");

  // 制御文字除去
  s = s.replace(/[\u0000-\u001F\u007F]/g, " ");

  return s;
}


// ----------------------------------------------------------
// ② UTF-16 セーフな 24 文字 × 2 行分割
// ----------------------------------------------------------
function splitByLengthSafe(str, size, lines = 2) {
  if (!str) return Array(lines).fill("");

  // UTF-16 サロゲートペア安全な配列化
  const chars = Array.from(str);

  const result = [];
  for (let i = 0; i < lines; i++) {
    const start = i * size;
    result[i] = chars.slice(start, start + size).join("");
  }
  return result;
}


// ----------------------------------------------------------
// ③ 正規化ヘルパ
// ----------------------------------------------------------
function norm(v) {
  return normalizeForSJIS(v ?? "");
}


// ==========================================================
// 🟥 メイン処理：ゆうパックCSV生成（BOM付きUTF-8）
// ==========================================================
async function convertToJapanPost(csvFile, sender) {
  console.log("📮 ゆうパック変換開始（完全版 BOM付きUTF-8）");

  const csvText = await csvFile.text();
  const rows = parseCsvSafe(csvText);
  const data = rows.slice(1);

  const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, "/");
  const output = [];


  // =======================
  // ご依頼主（送付元）
  // =======================
  const sAddrRaw = norm(sender.address);
  const [sPref, sAfterPref] = splitAddressPref(sAddrRaw);
  const [sCity, sAfterCity] = splitCity(sAfterPref);
  const sRest = splitByLengthSafe(sAfterCity, 24, 2);
  const senderAddrLines = [
    norm(sPref),
    norm(sCity),
    norm(sRest[0]),
    norm(sRest[1])
  ];

  const senderName   = norm(sender.name);
  const senderPostal = norm(sender.postal);
  const senderPhone  = norm(sender.phone);


  // =======================
  // 宛先（受取人）
  // =======================
  for (const r of data) {

    const name    = norm(r[12] || "");
    const postal  = norm(cleanTelPostal(r[10] || ""));
    const addrRaw = norm(r[11] || "");
    const phone   = norm(cleanTelPostal(r[13] || ""));
    const orderNo = norm(cleanOrderNumber(r[1] || ""));

    const [pref, afterPref] = splitAddressPref(addrRaw);
    const [city, afterCity] = splitCity(afterPref);

    const restLines = splitByLengthSafe(afterCity, 24, 2);
    const toAddrLines = [
      norm(pref),
      norm(city),
      norm(restLines[0]),
      norm(restLines[1])
    ];


    // =======================
    // ゆうパックCSV 1行生成
    // =======================
    const row = [];

    row.push("1", "0", "", "", "", "", "1");

    row.push(name);
    row.push("様");
    row.push("");

    row.push(postal);

    row.push(...toAddrLines);

    row.push(phone, "", "", "");
    row.push("", "", "");

    row.push(senderName, "", "", senderPostal);
    row.push(...senderAddrLines);

    row.push(senderPhone, "", orderNo, "");
    row.push("ブーケ加工品", "", "");

    row.push(todayStr, "", "", "", "", "");


    // 列数調整（ゆうパック仕様）
    while (row.length < 64) row.push("");
    row.push("0");
    while (row.length < 71) row.push("");
    row.push("0");

    output.push(row);
  }


  // ==========================================================
  // CSV（CRLF & ダブルクォート囲み）
  // ==========================================================
  const csvOut = output
    .map(r => r.map(v => `"${v}"`).join(","))
    .join("\r\n");


  // ==========================================================
  // 2. 【修正】Shift-JIS エンコードを BOM付きUTF-8 に変更
  // ==========================================================
  // const sjisArray = Encoding.convert(
  //   Encoding.stringToCode(csvOut),
  //   "SJIS"
  // );
  // return new Blob([new Uint8Array(sjisArray)], {
  //   type: "text/csv"
  // });
  
  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]); // BOM (Byte Order Mark)
  return new Blob([bom, csvOut], {
    type: "text/csv;charset=utf-8"
  });
}


// ==========================================================
// 🟩 佐川（e飛伝Ⅱ）ヘッダ付き・完全版
// 2. 【修正】文字化け対策のため、BOM付きUTF-8に変更
// ==========================================================
async function convertToSagawa(csvFile, sender) {
  console.log("📦 佐川変換開始（ヘッダ対応 完全版 BOM付きUTF-8）");

  const csvText = await csvFile.text();
  const rows = parseCsvSafe(csvText);
  const data = rows.slice(1);

  const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
  const output = [];

  // ==========================================================
  // 1行目：佐川 e飛伝Ⅱ 正式ヘッダ（全74項目）
  // ==========================================================
  const header = [
    "お届け先コード取得区分","お届け先コード","お届け先電話番号","お届け先郵便番号",
    "お届け先住所１","お届け先住所２","お届け先住所３",
    "お届け先名称１","お届け先名称２","お客様管理番号","お客様コード",
    "部署ご担当者コード取得区分","部署ご担当者コード","部署ご担当者名称",
    "荷送人電話番号","ご依頼主コード取得区分","ご依頼主コード",
    "ご依頼主電話番号","ご依頼主郵便番号","ご依頼主住所１",
    "ご依頼主住所２","ご依頼主名称１","ご依頼主名称２",
    "荷姿","品名１","品名２","品名３","品名４","品名５",
    "荷札荷姿","荷札品名１","荷札品名２","荷札品名３","荷札品名４","荷札品名５",
    "荷札品名６","荷札品名７","荷札品名８","荷札品名９","荷札品名１０","荷札品名１１",
    "出荷個数","スピード指定","クール便指定","配達日",
    "配達指定時間帯","配達指定時間（時分）","代引金額","消費税","決済種別","保険金額",
    "指定シール１","指定シール２","指定シール３",
    "営業所受取","SRC区分","営業所受取営業所コード","元着区分",
    "メールアドレス","ご不在時連絡先","出荷日","お問い合せ送り状No.",
    "出荷場印字区分","集約解除指定","編集01","編集02","編集03","編集04",
    "編集05","編集06","編集07","編集08","編集09","編集10"
  ];

  // 74列に満たない場合は埋める（形式上の保険）
  const headerRow = Array(74).fill("");
  header.forEach((v, i) => headerRow[i] = v);
  output.push(headerRow);

  // ==========================================================
  // ご依頼主情報（発送元）
  // ==========================================================
  const senderAddrLines = splitByLength(sender.address, 25, 2);

  // ==========================================================
  // データ本体
  // ==========================================================
  for (const r of data) {
    const out = Array(74).fill("");

    const addrFull = r[11] || "";
    const postal   = cleanTelPostal(r[10] || "");
    const tel      = cleanTelPostal(r[13] || "");
    const name     = r[12] || "";
    const orderNo  = cleanOrderNumber(r[1] || "");

    const toAddrLines = splitByLength(addrFull, 25, 3);

    // ================================
    // e飛伝Ⅱ 仕様に合わせたマッピング
    // ================================

    out[0] = "";              // お届け先コード取得区分
    out[1] = "";              // お届け先コード
    out[2] = tel;             // お届け先電話番号
    out[3] = postal;          // お届け先郵便番号

    out[4] = toAddrLines[0];  // お届け先住所1
    out[5] = toAddrLines[1];  // お届け先住所2
    out[6] = toAddrLines[2];  // お届け先住所3

    out[7] = name;            // お届け先名称1
    out[8] = "";              // お届け先名称2
    out[9] = orderNo;         // お客様管理番号
    out[10] = "";             // お客様コード

    out[11] = "";             // 部署ご担当者コード取得区分
    out[12] = "";             // 部署ご担当者コード
    out[13] = "";             // 部署ご担当者名称

    out[14] = sender.phone;   // 荷送人電話番号
    out[15] = "";             // ご依頼主コード取得区分
    out[16] = "";             // ご依頼主コード
    out[17] = sender.phone;   // ご依頼主電話番号
    out[18] = sender.postal;  // ご依頼主郵便番号

    out[19] = senderAddrLines[0]; // ご依頼主住所1
    out[20] = senderAddrLines[1]; // ご依頼主住所2
    out[21] = sender.name;        // ご依頼主名称1
    out[22] = "";                 // ご依頼主名称2

    out[23] = "";                 // 荷姿
    out[24] = "ブーケ加工品";     // 品名1
    out[25] = ""; out[26] = ""; out[27] = ""; out[28] = ""; // 品名2〜5

    // 荷札関係（未使用）
    for (let i = 29; i <= 39; i++) out[i] = "";

    out[40] = "1";               // 出荷個数（1固定）
    out[41] = "";                // スピード指定
    out[42] = "";                // クール便指定
    out[43] = "";                // 配達日

    out[44] = "";                // 配達時間帯
    out[45] = "";                // 指定時刻
    out[46] = "";                // 代引金額
    out[47] = "";                // 消費税
    out[48] = "";                // 決済種別
    out[49] = "";                // 保険金額

    out[50] = ""; out[51] = ""; out[52] = ""; // 指定シール1〜3
    out[53] = "";                // 営業所受取
    out[54] = "";                // SRC区分
    out[55] = "";                // 営業所コード
    out[56] = "";                // 元着区分
    out[57] = "";                // メール
    out[58] = "";                // ご不在連絡先
    out[59] = todayStr;          // 出荷日
    out[60] = "";                // 問合せ送り状No.

    out[61] = "";                // 出荷場印字区分
    out[62] = "";                // 集約解除指定

    // 編集01〜編集10
    for (let i = 63; i < 74; i++) out[i] = "";

    output.push(out);
  }

  // ==========================================================
  // CSV 出力
  // ==========================================================
  const csvTextOut = output
    .map(r => r.map(v => `"${v ?? ""}"`).join(","))
    .join("\r\n");

  // 2. 【修正】Shift-JIS エンコードを BOM付きUTF-8 に変更
  // const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut), "SJIS");
  // return new Blob([new Uint8Array(sjis)], { type: "text/csv" });

  const bom = new Uint8Array([0xEF, 0xBB, 0xBF]); // BOM (Byte Order Mark)
  return new Blob([bom, csvTextOut], { type: "text/csv;charset=utf-8" });
}


  // ==========================================================
  // 🟥 ヤマト（住所1列・Excel）
  // 1. 【修正】CSV読み込みを parseCsvSafe に変更
  // ==========================================================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚【テスト】ヤマト開始（住所1列）");

    const csvText = await csvFile.text();
    // ↓ 修正: parseCsvSafe を使用して、安全にCSVを読み込む
    const rows    = parseCsvSafe(csvText);
    const data    = rows.slice(1);

    const res = await fetch("./js/newb2web_template1.xlsx");
    const wb = XLSX.read(await res.arrayBuffer(),{type:"array"});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const header = XLSX.utils.sheet_to_json(sheet,{header:1})[0];

    function colLetter(i){
      let s=""; while(i>=0){ s=String.fromCharCode(i%26+65)+s; i=Math.floor(i/26)-1; }
      return s;
    }
    function idx(key){
      return header.findIndex(h=>typeof h==="string"&&h.includes(key));
    }

    const map = {
      order : idx("お客様管理番号"),
      type  : idx("送り状種類"),
      cool  : idx("クール区分"),
      ship1 : idx("出荷予定日"),
      ship2 : idx("出荷日"),
      tel   : idx("お届け先電話番号"),
      zip   : idx("お届け先郵便番号"),
      adr   : idx("お届け先住所"),
      apt   : idx("お届け先アパート"),
      name  : idx("お届け先名"),
      honor : idx("敬称"),
      sTel  : idx("ご依頼主電話番号"),
      sZip  : idx("ご依頼主郵便番号"),
      sAdr  : idx("ご依頼主住所"),
      sApt  : idx("ご依頼主アパート"),
      sName : idx("ご依頼主名"),
      item  : idx("品名１")
    };

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    let rowExcel = 2;

    function set(i,val){
      if(i < 0) return;
      sheet[colLetter(i)+rowExcel] = { v: val, t: "s" };
    }

    for(const r of data){
      // r[1]: 注文番号, r[14]: 電話番号, r[11]: 郵便番号, r[13]: 氏名, r[12]: 住所
      // 読み込みが parseCsvSafe になったため、インデックスは元の入力CSVの構造に合わせる
      const order = cleanOrderNumber(r[1]||"");
      const tel   = cleanTelPostal(r[14]||"");
      const zip   = cleanTelPostal(r[11]||"");
      const name  = r[13]||"";
      const adr   = r[12]||"";  // ★住所1列

      set(map.order, order);
      set(map.type, "0");
      set(map.cool, "0");
      set(map.ship1, todayStr);
      set(map.ship2, todayStr);

      set(map.tel, tel);
      set(map.zip, zip);

      set(map.adr, adr);
      set(map.apt, "");

      set(map.name, name);
      set(map.honor, "様");

      set(map.sTel, sender.phone);
      set(map.sZip, sender.postal);
      set(map.sAdr, sender.address);
      set(map.sApt, "");
      set(map.sName, sender.name);

      set(map.item, "ブーケ加工品");
      // map.item2 が定義されていないため、この行は元のコードのままとするが、
      // Excelテンプレートに "品名2" があれば、そのインデックスを使用すべき
      // set(map.item2, order);    
      

      rowExcel++;
    }

    return wb;
  }

  // ============================
  // 変換ボタン
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file    = fileInput.files[0];
      const courier = courierSelect.value;
      if (!file) return;

      const sender = getSenderInfo();
      showLoading(true);

      try {
        if (courier === "yamato") {
          mergedWorkbook = await convertToYamato(file, sender);
          convertedCSV   = null;
        } else if (courier === "japanpost") {
          convertedCSV   = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
        } else {
          convertedCSV   = await convertToSagawa(file, sender);
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
          // 2. 【修正】ファイル名に utf8 を含める（任意）
          courier === "japanpost" ? "yupack_import_utf8.csv" :
          courier === "sagawa"    ? "sagawa_import_utf8.csv" :
          "output_utf8.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = name;
        link.click();
        URL.revokeObjectURL(link.href);
      }
    });
  }

})();
