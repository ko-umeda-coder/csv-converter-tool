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

  // 住所分割：都道府県 / 市区郡町村 / 丁番地・その他 / 建物名
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
    let rest = pref ? address.slice(pref.length) : address;

    const [city, ...after] = rest.split(/(?<=市|区|町|村)/);
    rest = after.join("");

    let building = "";
    const bMatch = rest.match(/(ビル|マンション|ハイツ|アパート|号室|F|階).*/);
    if (bMatch) {
      building = bMatch[0];
      rest = rest.replace(building, "");
    }

    return {
      pref,
      city: city || "",
      rest: rest || "",
      building: building || ""
    };
  }

 // ============================
// ヤマト B2クラウド 変換（95列・ヘッダ名ベース）
// ============================
async function convertToYamato(csvFile, sender) {
  console.log("🚚 ヤマトB2変換開始");

  // 入力CSV読み込み
  const csvText = await csvFile.text();
  const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
  const data    = rows.slice(1); // 1行目ヘッダを除外

  // テンプレート読込（ヤマト正解.xlsx と同じ構成の newb2web_template1.xlsx）
  const res = await fetch("./js/newb2web_template1.xlsx");
  const buf = await res.arrayBuffer();
  const wb  = XLSX.read(buf, { type: "array" });

  // 最初のシートを使用（≒「外部データ取り込み基本レイアウト」）
  const sheetName = wb.SheetNames[0];
  const sheet     = wb.Sheets[sheetName];

  // 1行目ヘッダ行を配列で取得
  const headerRows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headerRow  = headerRows[0] || [];

  // ヘッダ内で「～を含む」列インデックスを検索
  function findHeaderIndex(keyword) {
    const idx = headerRow.findIndex(h =>
      typeof h === "string" && h.includes(keyword)
    );
    if (idx === -1) {
      console.warn("⚠ ヘッダが見つかりません:", keyword);
    }
    return idx;
  }

  // Excel列番号 → 列名（0=A, 1=B,...）
  function colLetter(idx) {
    let s = "";
    let n = idx;
    while (n >= 0) {
      s = String.fromCharCode((n % 26) + 65) + s;
      n = Math.floor(n / 26) - 1;
    }
    return s;
  }

  // 使うヘッダのマッピング定義（keyword はセル内に含まれている文字）
  const ruleDefs = [
    // お客様管理番号 = CSV B列（ご注文番号をクレンジング）
    { keyword: "お客様管理番号",   type: "csv",   col: 1,  clean: "order" },

    // 固定値
    { keyword: "送り状種類",       type: "value", value: "0" },
    { keyword: "クール区分",       type: "value", value: "0" },

    // 出荷予定日・出荷日 = TODAY
    { keyword: "出荷予定日",       type: "today" },
    { keyword: "出荷日",           type: "today" }, // シートにあればセット

    // お届け先（CSV側：K=10, L=11, M=12, N=13）
    { keyword: "お届け先電話番号", type: "csv",   col: 13, clean: "tel" },
    { keyword: "お届け先郵便番号", type: "csv",   col: 10, clean: "postal" },
    { keyword: "お届け先住所",     type: "csv",   col: 11 },
    { keyword: "お届け先アパートマンション", type: "csv", col: 11 },
    { keyword: "お届け先名",       type: "csv",   col: 12 },
    { keyword: "敬称",             type: "value", value: "様" },

    // ご依頼主（UI入力）
    { keyword: "ご依頼主電話番号",           type: "sender", field: "phone" },
    { keyword: "ご依頼主郵便番号",           type: "sender", field: "postal" },
    { keyword: "ご依頼主住所",               type: "sender", field: "address" },
    { keyword: "ご依頼主アパートマンション", type: "sender", field: "address" },
    { keyword: "ご依頼主名",                 type: "sender", field: "name" },

    // 品名１ 固定
    { keyword: "品名１",           type: "value", value: "ブーケ加工品" },
  ];

  // 1回だけヘッダ→列インデックスを解決してキャッシュ
  const headerIndexMap = {};
  for (const rule of ruleDefs) {
    const idx = findHeaderIndex(rule.keyword);
    if (idx >= 0) {
      headerIndexMap[rule.keyword] = idx;
    }
  }

  // 日付文字列
  const today = new Date();
  const todayStr =
    `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

  // 2行目から順に書き込み
  let rowExcel = 2;

  for (const r of data) {
    for (const rule of ruleDefs) {
      const colIdx = headerIndexMap[rule.keyword];
      if (colIdx === undefined) continue; // 該当ヘッダがテンプレートに無ければスキップ

      const col = colLetter(colIdx);
      const cellRef = col + rowExcel;

      let v = "";

      switch (rule.type) {
        case "value":
          v = rule.value;
          break;

        case "today":
          v = todayStr;
          break;

        case "csv": {
          const src = r[rule.col] || "";
          if (rule.clean === "tel" || rule.clean === "postal") {
            v = cleanTelPostal(src);
          } else if (rule.clean === "order") {
            v = cleanOrderNumber(src);
          } else {
            v = src;
          }
          break;
        }

        case "sender": {
          const val = sender[rule.field] || "";
          if (rule.field === "phone" || rule.field === "postal") {
            v = cleanTelPostal(val);
          } else {
            v = val;
          }
          break;
        }
      }

      sheet[cellRef] = { v, t: "s" };
    }

    rowExcel++;
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

  for (const r of data) {
    const name        = r[12] || "";                     // M列：氏名
    const postal      = cleanTelPostal(r[10] || "");     // K列：郵便番号
    const addressFull = r[11] || "";                     // L列：住所
    const phone       = cleanTelPostal(r[13] || "");     // N列：電話番号
    const orderNo     = cleanOrderNumber(r[1] || "");    // B列：ご依頼主部署名として使用

    const addr = splitAddress(addressFull);
    const sendAddr = splitAddress(sender.address);

    const row = [];

    // 👉 ここから 72 列固定で push
    row.push("1");           // 1 商品
    row.push("0");           // 2 着払/代引
    row.push("");            // 3
    row.push("");            // 4
    row.push("");            // 5
    row.push("");            // 6
    row.push("1");           // 7 作成数

    row.push(name);          // 8 お届け先のお名前
    row.push("様");          // 9 お届け先の敬称
    row.push("");            // 10 お名前（カナ）
    row.push(postal);        // 11 郵便番号
    row.push(addr.pref);     // 12 都道府県
    row.push(addr.city);     // 13 市区町村郡
    row.push(addr.rest);     // 14 丁目番地号
    row.push(addr.building); // 15 建物
    row.push(phone);         // 16 電話番号
    row.push("");            // 17 法人名
    row.push("");            // 18 部署
    row.push("");            // 19 メール

    row.push("");            // 20 空港略称
    row.push("");            // 21 空港コード
    row.push("");            // 22 受取人様のお名前

    row.push(sender.name);           // 23 ご依頼主名
    row.push("");                    // 24 敬称
    row.push("");                    // 25 カナ
    row.push(sender.postal);         // 26 郵便番号
    row.push(sendAddr.pref);         // 27 都道府県
    row.push(sendAddr.city);         // 28 市区町村
    row.push(sendAddr.rest);         // 29 丁番地
    row.push(sendAddr.building);     // 30 建物
    row.push(sender.phone);          // 31 電話番号

    row.push("");                    // 32 法人名
    row.push(orderNo);               // 33 部署名 ←ここに注文番号
    row.push("");                    // 34 メール

    row.push("ブーケ加工品");        // 35 品名
    row.push("");                    // 36 品名番号
    row.push("");                    // 37 個数

    row.push(todayStr);             // 38 発送予定日
    row.push("");                   // 39 発送予定時間帯
    row.push("");                   // 40 セキュリティ
    row.push("");                   // 41 重量
    row.push("");                   // 42 損害要償額
    row.push("");                   // 43 保冷

    row.push("");                   // 44 こわれもの
    row.push("");                   // 45 なまもの
    row.push("");                   // 46 ビン類
    row.push("");                   // 47 逆さま厳禁
    row.push("");                   // 48 下積み厳禁

    row.push("");                   // 49 予備
    row.push("");                   // 50 差出予定日
    row.push("");                   // 51 差出予定時間帯
    row.push("");                   // 52 配達希望日
    row.push("");                   // 53 配達希望時間帯
    row.push("");                   // 54 クラブ本数
    row.push("");                   // 55 使用日
    row.push("");                   // 56 使用時間
    row.push("");                   // 57 搭乗日
    row.push("");                   // 58 搭乗時間
    row.push("");                   // 59 搭乗便名
    row.push("");                   // 60 復路発送予定日
    row.push("");                   // 61 支払方法
    row.push("");                   // 62 摘要
    row.push("");                   // 63 サイズ
    row.push("");                   // 64 差出方法
    row.push("0");                  // 65 割引
    row.push("");                   // 66 代引金額
    row.push("");                   // 67 消費税
    row.push("");                   // 68 配達予定通知
    row.push("");                   // 69 配達完了通知
    row.push("");                   // 70 不在通知
    row.push("");                   // 71 郵便局留通知
    row.push("0");                  // 72 配達完了通知(依頼主)

    output.push(row);
  }

  // 👉 ヘッダなし・72列の CSV 出力
  const csvOut = output.map(row => row.map(v => `"${v}"`).join(",")).join("\r\n");
  const sjis = Encoding.convert(Encoding.stringToCode(csvOut), "SJIS");

  return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
}

  
  // ============================
// 佐川 e飛伝Ⅱ（74列固定・住所分割対応）
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

  // ◆ ご依頼主住所分割
  const sendAddr = splitAddress(sender.address);
  const sendAddr1 = (sendAddr.pref || "") + (sendAddr.city || ""); // 都道府県 + 市区町村郡
  const sendAddr2 = ((sendAddr.rest || "") + (sendAddr.building || "")).trim(); // 丁目番地号 + 建物名

  for (const r of data) {
    const out = Array(headers.length).fill("");

    const orderNumber = cleanOrderNumber(r[1] || "");
    const postal      = cleanTelPostal(r[10] || "");
    const addressFull = r[11] || "";
    const name        = r[12] || "";
    const phone       = cleanTelPostal(r[13] || "");
    const addr        = splitAddress(addressFull);

    // ======== ★ 各列へのセット（正解仕様） ========
    out[0]  = "0";                      // A: コード取得区分
    out[2]  = phone;                    // C: 電話番号
    out[3]  = postal;                   // D: 郵便番号
    out[4]  = addr.pref + addr.city;    // E: 住所1
    out[5]  = addr.rest;                // F: 住所2
    out[6]  = addr.building;            // G: 住所3
    out[7]  = name;                     // H: 名称1
    out[8]  = orderNumber;              // I: 名称2 ← 注文番号

    out[17] = sender.phone;             // R: ご依頼主電話番号
    out[18] = sender.postal;            // S: ご依頼主郵便番号

    // ⭐修正：住所1 / 住所2 を分割してセット
    out[19] = sendAddr1;                // T: ご依頼主住所１（都道府県＋市区町村）
    out[20] = sendAddr2;                // U: ご依頼主住所２（丁目番地号＋建物名）

    out[21] = sender.name;              // V: ご依頼主名称１
    out[25] = "ブーケ加工品";          // Z: 品名１
    out[58] = todayStr;                 // BG: 出荷日

    output.push(out);
  }

  // CSV書き出し（ヘッダあり）
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
