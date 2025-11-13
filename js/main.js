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
  console.log("🚚【テスト】ヤマトB2（住所なし）");

  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1);

  const res = await fetch("./js/newb2web_template1.xlsx");
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type: "array" });

  const sheetName = wb.SheetNames[0];
  const sheet     = wb.Sheets[sheetName];
  const headerRow = XLSX.utils.sheet_to_json(sheet, {header:1})[0];

  function colLetter(idx){ …同じ… }

  const index = keyword =>
    headerRow.findIndex(h => typeof h === "string" && h.includes(keyword));

  const map = {
    customer: index("お客様管理番号"),
    type: index("送り状種類"),
    cool: index("クール区分"),
    shipdate: index("出荷予定日"),
    deltel: index("お届け先電話番号"),
    delzip: index("お届け先郵便番号"),
    deladdr: index("お届け先住所"),
    delapt: index("お届け先アパート"),
    delname: index("お届け先名"),
    honor: index("敬称"),
    sndtel: index("ご依頼主電話番号"),
    sndzip: index("ご依頼主郵便番号"),
    sndaddr: index("ご依頼主住所"),
    sndapt: index("ご依頼主アパートマンション"),
    sndname: index("ご依頼主名"),
    item: index("品名１"),
    shipdate2: index("出荷日"),
  };

  const today = new Date();
  const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  let rowExcel = 2;

  for (const r of data) {
    const o = cleanOrderNumber(r[1] || "");
    const tel = cleanTelPostal(r[13] || "");
    const zip = cleanTelPostal(r[10] || "");
    const name = r[12] || "";

    function set(colIdx, val) {
      if (colIdx < 0) return;
      const cell = colLetter(colIdx) + rowExcel;
      sheet[cell] = { v: val, t:"s" };
    }

    set(map.customer, o);
    set(map.type, "0");
    set(map.cool, "0");
    set(map.shipdate, todayStr);
    set(map.shipdate2, todayStr);

    set(map.deltel, tel);
    set(map.delzip, zip);

    // ★住所を完全に空欄にする
    set(map.deladdr, "");
    set(map.delapt, "");

    set(map.delname, name);
    set(map.honor, "様");

    set(map.sndtel, sender.phone);
    set(map.sndzip, sender.postal);

    // ★依頼主住所も空欄
    set(map.sndaddr, "");
    set(map.sndapt, "");

    set(map.sndname, sender.name);
    set(map.item, "ブーケ加工品");

    rowExcel++;
  }

  return wb;
}



async function convertToJapanPost(csvFile, sender) {
  console.log("📮【テスト】ゆうパック（住所なし）");

  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1);

  const output = [];
  const today  = new Date();
  const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  for (const r of data) {
    const name   = r[12] || "";
    const postal = cleanTelPostal(r[10] || "");
    const phone  = cleanTelPostal(r[13] || "");
    const orderNo = cleanOrderNumber(r[1] || "");

    const row = [];

    row.push("1"); // 商品
    row.push("0"); // 着払/代引
    row.push(""); row.push(""); row.push(""); row.push(""); 
    row.push("1"); // 作成数

    row.push(name);  // お届け先名
    row.push("様");
    row.push(""); // カナ
    row.push(postal);

    // ======= ★ 住所系すべて空欄にする =======
    row.push(""); // 都道府県
    row.push(""); // 市区町村
    row.push(""); // 丁番地
    row.push(""); // 建物

    row.push(phone);
    row.push(""); row.push(""); row.push("");

    // 空港など
    row.push(""); row.push(""); row.push("");

    // ご依頼主
    row.push(sender.name);
    row.push(""); row.push("");
    row.push(sender.postal);

    // ★住所なし
    row.push(""); 
    row.push("");
    row.push("");
    row.push("");

    row.push(sender.phone);

    row.push(""); // 法人
    row.push(orderNo); // 部署名に注文番号
    row.push(""); // メール

    row.push("ブーケ加工品");
    row.push(""); row.push("");

    row.push(todayStr); // 発送予定日
    row.push(""); row.push(""); row.push(""); row.push(""); row.push("");

    // 注意書き・その他すべて空欄
    for (let i = 0; i < (72 - row.length - 1); i++) row.push("");

    row.push("0"); // 最後の列（配達完了通知 依頼主）

    output.push(row);
  }

  const csvOut = output.map(row => row.map(v => `"${v}"`).join(",")).join("\r\n");
  const sjis = Encoding.convert(Encoding.stringToCode(csvOut), "SJIS");
  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
}


  
async function convertToSagawa(csvFile, sender) {
  console.log("📦【テスト】佐川（住所なし）");

  const headers = [ ... 同じ 74項目 ... ];

  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1);

  const today = new Date();
  const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  const output = [];

  for (const r of data) {
    const out = Array(headers.length).fill("");

    const orderNo = cleanOrderNumber(r[1] || "");
    const postal  = cleanTelPostal(r[10] || "");
    const name    = r[12] || "";
    const phone   = cleanTelPostal(r[13] || "");

    out[0]  = "0";
    out[2]  = phone;
    out[3]  = postal;

    // ======= ★ 住所1/2/3 全部空欄 =======
    out[4] = ""; 
    out[5] = "";
    out[6] = "";

    out[7] = name;
    out[8] = orderNo;

    out[17] = sender.phone;
    out[18] = sender.postal;

    // ★ご依頼主住所も空
    out[19] = "";
    out[20] = "";

    out[21] = sender.name;

    out[25] = "ブーケ加工品";
    out[58] = todayStr;

    output.push(out);
  }

  const csvTextOut = [
    headers.join(","),
    ...output.map(r => r.map(v => `"${v}"`).join(","))
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
