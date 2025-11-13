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
  // ヤマト B2 変換
  // ============================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚 ヤマト変換開始");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
    const data    = rows.slice(1); // 1行目ヘッダを除外

    // テンプレート読込
    const res = await fetch("./js/newb2web_template1.xlsx");
    const buf = await res.arrayBuffer();
    const wb  = XLSX.read(buf, { type: "array" });

    const sheetName = wb.SheetNames[0];
    const sheet     = wb.Sheets[sheetName];

    // シートの1行目からヘッダ文言を取得
    const headerRow = XLSX.utils.sheet_to_json(sheet, { header: 1 })[0] || [];
    const headerMap = {}; // "お届け先電話番号" → 列インデックス
    headerRow.forEach((h, idx) => {
      if (typeof h === "string" && h.trim()) {
        headerMap[h.trim()] = idx;
      }
    });

    // 列番号 → Excel列文字（0:A, 1:B, ...）
    function colLetter(idx) {
      let s = "";
      let n = idx;
      while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
      }
      return s;
    }

    // マッピング定義（ヘッダ名ベース）
    const mapping = [
      { header: "お客様管理番号",     type: "csv",    col: 1,  clean: "order" },
      { header: "送り状種類",         type: "value",  value: "0" },
      { header: "クール区分",         type: "value",  value: "0" },
      { header: "出荷予定日",         type: "today" },
      { header: "お届け先電話番号",   type: "csv",    col: 13, clean: "tel" },
      { header: "お届け先郵便番号",   type: "csv",    col: 10, clean: "postal" },
      { header: "お届け先住所",       type: "csv",    col: 11 },
      { header: "お届け先アパートマンション名", type: "csv", col: 11 },
      { header: "お届け先名",         type: "csv",    col: 12 },
      { header: "敬称",               type: "value",  value: "様" },
      { header: "ご依頼主電話番号",   type: "sender", field: "phone" },
      { header: "ご依頼主郵便番号",   type: "sender", field: "postal" },
      { header: "ご依頼主住所",       type: "sender", field: "address" },
      { header: "ご依頼主アパートマンション", type: "sender", field: "address" },
      { header: "ご依頼主名",         type: "sender", field: "name" },
      { header: "品名１",             type: "value",  value: "ブーケ加工品" },
    ];

    const today = new Date();
    const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

    let rowExcel = 2; // Excel上の2行目からデータ

    for (const r of data) {
      for (const rule of mapping) {
        const idx = headerMap[rule.header];
        if (idx === undefined) continue; // テンプレートにそのヘッダが無い場合はスキップ
        const col = colLetter(idx);
        const cellRef = col + rowExcel;

        let v = "";
        if (rule.type === "value") {
          v = rule.value;
        } else if (rule.type === "today") {
          v = todayStr;
        } else if (rule.type === "csv") {
          const src = r[rule.col] || "";
          if (rule.clean === "tel" || rule.clean === "postal") {
            v = cleanTelPostal(src);
          } else if (rule.clean === "order") {
            v = cleanOrderNumber(src);
          } else {
            v = src;
          }
        } else if (rule.type === "sender") {
          v = sender[rule.field] || "";
        }

        sheet[cellRef] = { v: v, t: "s" };
      }
      rowExcel++;
    }

    return wb;
  }

  // ============================
  // ゆうパック（ゆうプリR）変換
  // ============================
  async function convertToJapanPost(csvFile, sender) {
    console.log("📮 ゆうパック変換開始");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
    const data    = rows.slice(1);

    const output = [];
    const today  = new Date();
    const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

    const senderAddr = splitAddress(sender.address);

    for (const r of data) {
      const orderNumber = cleanOrderNumber(r[1] || "");
      const name        = r[12] || "";
      const postal      = cleanTelPostal(r[10] || "");
      const addressFull = r[11] || "";
      const phone       = cleanTelPostal(r[13] || "");
      const addr        = splitAddress(addressFull);

      const rowOut = [];

      // ※列順はユーザー指定通り
      rowOut.push("1");             // 商品
      rowOut.push("0");             // 着払/代引
      rowOut.push("");              // ゴルフ/スキー/空港
      rowOut.push("");              // 往復
      rowOut.push("");              // 書留/特定記録
      rowOut.push("");              // 配達方法
      rowOut.push("1");             // 作成数

      rowOut.push(name);            // お届け先のお名前
      rowOut.push("様");            // 敬称
      rowOut.push("");              // お届け先カナ
      rowOut.push(postal);          // 郵便番号
      rowOut.push(addr.pref);       // 都道府県
      rowOut.push(addr.city);       // 市区町村郡
      rowOut.push(addr.rest);       // 丁目番地号
      rowOut.push(addr.building);   // 建物名・部屋番号など
      rowOut.push(phone);           // 電話番号
      rowOut.push("");              // 法人名
      rowOut.push("");              // 部署名
      rowOut.push("");              // メールアドレス

      rowOut.push("");              // 空港略称
      rowOut.push("");              // 空港コード
      rowOut.push("");              // 受取人様のお名前

      rowOut.push(sender.name);           // ご依頼主のお名前
      rowOut.push("");                    // ご依頼主の敬称
      rowOut.push("");                    // ご依頼主カナ
      rowOut.push(sender.postal);         // ご依頼主郵便番号

      rowOut.push(senderAddr.pref);       // ご依頼主都道府県
      rowOut.push(senderAddr.city);       // ご依頼主市区町村郡
      rowOut.push(senderAddr.rest);       // ご依頼主丁目番地号
      rowOut.push(senderAddr.building);   // ご依頼主建物名・部屋番号
      rowOut.push(sender.phone);          // ご依頼主電話番号

      rowOut.push("");                    // ご依頼主法人名
      rowOut.push(orderNumber);           // ご依頼主部署名（ここにご注文番号）
      rowOut.push("");                    // ご依頼主メールアドレス

      rowOut.push("ブーケ加工品");        // 品名
      rowOut.push("");                    // 品名番号
      rowOut.push("");                    // 個数

      rowOut.push(todayStr);             // 発送予定日
      rowOut.push("");                   // 発送予定時間帯
      rowOut.push("");                   // セキュリティ
      rowOut.push("");                   // 重量
      rowOut.push("");                   // 損害要償額
      rowOut.push("");                   // 保冷

      rowOut.push("");                   // こわれもの
      rowOut.push("");                   // なまもの
      rowOut.push("");                   // ビン類
      rowOut.push("");                   // 逆さま厳禁
      rowOut.push("");                   // 下積み厳禁

      rowOut.push("");                   // 予備
      rowOut.push("");                   // 差出予定日
      rowOut.push("");                   // 差出予定時間帯
      rowOut.push("");                   // 配達希望日
      rowOut.push("");                   // 配達希望時間帯
      rowOut.push("");                   // クラブ本数
      rowOut.push("");                   // ご使用日(プレー日)
      rowOut.push("");                   // ご使用時間
      rowOut.push("");                   // 搭乗日
      rowOut.push("");                   // 搭乗時間
      rowOut.push("");                   // 搭乗便名
      rowOut.push("");                   // 復路発送予定日
      rowOut.push("");                   // お支払方法
      rowOut.push("");                   // 摘要/記事
      rowOut.push("");                   // サイズ
      rowOut.push("");                   // 差出方法
      rowOut.push("0");                  // 割引
      rowOut.push("");                   // 代金引換金額
      rowOut.push("");                   // うち消費税等
      rowOut.push("");                   // 配達予定日通知(お届け先)
      rowOut.push("");                   // 配達完了通知(お届け先)
      rowOut.push("");                   // 不在持戻り通知(お届け先)
      rowOut.push("");                   // 郵便局留通知(お届け先)
      rowOut.push("0");                  // 配達完了通知(依頼主)

      output.push(rowOut);
    }

    // ヘッダなしでCSV化
    const csvTextOut = output
      .map(row => row.map(v => `"${v ?? ""}"`).join(","))
      .join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv" });
  }

  // ============================
  // 佐川 e飛伝Ⅱ 変換
  // ============================
  async function convertToSagawa(csvFile, sender) {
    console.log("📦 佐川変換開始");

    // 公式ヘッダ順（A列〜）
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

    const output  = [];
    const today   = new Date();
    const todayStr = `${today.getFullYear()}/${String(today.getMonth()+1).padStart(2,"0")}/${String(today.getDate()).padStart(2,"0")}`;

    const senderAddr = splitAddress(sender.address);

    for (const r of data) {
      const out = Array(headers.length).fill("");

      const orderNumber = cleanOrderNumber(r[1] || "");
      const postal      = cleanTelPostal(r[10] || "");
      const addressFull = r[11] || "";
      const name        = r[12] || "";
      const phone       = cleanTelPostal(r[13] || "");
      const addr        = splitAddress(addressFull);

      // 列マッピング（ユーザー指定に基づく）
      out[0]  = "0";               // A: お届け先コード取得区分
      out[2]  = phone;             // C: お届け先電話番号
      out[3]  = postal;            // D: お届け先郵便番号
      out[4]  = addr.pref + addr.city; // E: お届け先住所１
      out[5]  = addr.rest;         // F: お届け先住所２
      out[6]  = addr.building;     // G: お届け先住所３
      out[7]  = name;              // H: お届け先名称１
      out[8]  = orderNumber;       // I: お届け先名称２（ご注文番号）

      out[17] = sender.phone;      // R: ご依頼主電話番号
      out[18] = sender.postal;     // S: ご依頼主郵便番号
      out[19] = sender.address;    // T: ご依頼主住所１
      out[20] = sender.address;    // U: ご依頼主住所２
      out[21] = sender.name;       // V: ご依頼主名称１

      out[25] = "ブーケ加工品";   // Z: 品名１
      out[58] = todayStr;          // BG: 出荷日

      output.push(out);
    }

    const csvTextOut = [
      headers.join(","),
      ...output.map(row => row.map(v => `"${v ?? ""}"`).join(","))
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
