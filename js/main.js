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
// メイン処理
// ============================
(async () => {
  await waitForXLSX();
  console.log("🔥 main.js 起動（住所25文字分割版）");

  const fileInput     = document.getElementById("csvFile");
  const fileWrapper   = document.getElementById("fileWrapper");
  const fileName      = document.getElementById("fileName");
  const convertBtn    = document.getElementById("convertBtn");
  const downloadBtn   = document.getElementById("downloadBtn");
  const messageBox    = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;
  let convertedCSV   = null;

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
      name:    document.getElementById("senderName").value.trim(),
      postal:  cleanTelPostal(document.getElementById("senderPostal").value.trim()),
      address: document.getElementById("senderAddress").value.trim(), // ← 1列入力を25文字分割で使用
      phone:   cleanTelPostal(document.getElementById("senderPhone").value.trim()),
    };
  }

  // ============================
  // UIまわり
  // ============================
  function setupFileInput() {
    fileInput.addEventListener("change", () => {
      if (fileInput.files.length > 0) {
        fileName.textContent = fileInput.files[0].name;
        fileWrapper.classList.add("has-file");
        convertBtn.disabled = false;
      } else {
        fileName.textContent = "";
        fileWrapper.classList.remove("has-file");
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
    return String(v).replace(/[^0-9\-]/g, "");
  }

  function cleanOrderNumber(v) {
    if (!v) return "";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "");
  }

  // 25文字ごとに分割するユーティリティ
  // maxParts で必要な行数を指定（足りない分は "" を返す）
  function splitByLength(text, partLen, maxParts) {
    const s = text || "";
    const parts = [];
    for (let i = 0; i < maxParts; i++) {
      const start = i * partLen;
      if (start >= s.length) {
        parts.push("");
      } else {
        parts.push(s.slice(start, start + partLen));
      }
    }
    return parts;
  }

  // ==========================================================
  // 🟦 ゆうパック（住所を25文字で最大4分割／72列固定）
  // ==========================================================
  async function convertToJapanPost(csvFile, sender) {
    console.log("📮 ゆうパック変換開始（住所25文字分割）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data    = rows.slice(1);

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    const output = [];

    // 送り主住所を4行まで25文字分割
    const senderAddrLines = splitByLength(sender.address, 25, 4);

    for (const r of data) {
      // ★ インポート元の列指定はテスト版から変更しない
      const name    = r[13] || "";                 // 宛名
      const postal  = cleanTelPostal(r[11] || ""); // 郵便番号
      const addrRaw = r[12] || "";                 // フル住所1列
      const phone   = cleanTelPostal(r[14] || ""); // 電話番号
      const orderNo = cleanOrderNumber(r[1] || "");// ご注文番号

      // お届け先住所を最大4行まで 25文字分割
      const toAddrLines = splitByLength(addrRaw, 25, 4);

      const row = [];

      // 1〜7
      row.push("1","0","","","","","1");

      // 8〜11
      row.push(name);      // 8 お届け先の名前
      row.push("様");      // 9 敬称
      row.push("");        // 10 カナ
      row.push(postal);    // 11 郵便番号

      // 12〜15 住所4行（25文字分割）
      row.push(toAddrLines[0]); // 12
      row.push(toAddrLines[1]); // 13
      row.push(toAddrLines[2]); // 14
      row.push(toAddrLines[3]); // 15

      // 16〜19
      row.push(phone);     // 16 電話
      row.push("");        // 17 法人名
      row.push("");        // 18 部署名
      row.push("");        // 19 メール

      // 20〜22（空港関連など）空欄
      row.push("","", "");

      // 23〜26 ご依頼主
      row.push(sender.name);    // 23 ご依頼主名
      row.push("");             // 24 敬称
      row.push("");             // 25 カナ
      row.push(sender.postal);  // 26 郵便番号

      // 27〜30 ご依頼主住所（25文字×4）
      row.push(senderAddrLines[0]); // 27
      row.push(senderAddrLines[1]); // 28
      row.push(senderAddrLines[2]); // 29
      row.push(senderAddrLines[3]); // 30

      // 31〜34 ご依頼主電話・部署名など
      row.push(sender.phone); // 31 電話
      row.push("");           // 32 法人名
      row.push(orderNo);      // 33 部署名 ← ご注文番号
      row.push("");           // 34 ご依頼主メール

      // 35〜37 品名等
      row.push("ブーケ加工品"); // 35 品名
      row.push("");             // 36 品名番号
      row.push("");             // 37 個数

      // 38〜43 発送予定日など
      row.push(todayStr); // 38 発送予定日
      row.push("");       // 39
      row.push("");       // 40
      row.push("");       // 41
      row.push("");       // 42
      row.push("");       // 43

      // 44〜64 各種フラグ等 空欄
      while (row.length < 64) row.push("");

      // 65 割引
      row.push("0"); // 65 割引

      // 66〜71 空欄
      while (row.length < 71) row.push("");

      // 72 配達完了通知(依頼主)
      row.push("0");

      output.push(row);
    }

    const csvOut = output
      .map(r => r.map(v=>`"${v ?? ""}"`).join(","))
      .join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }

  // ==========================================================
  // 🟩 佐川（住所を25文字で分割／74列固定）
  // ==========================================================
  async function convertToSagawa(csvFile, sender) {
    console.log("📦 佐川変換開始（住所25文字分割）");

    const headers = [
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

    const csvText = await csvFile.text();
    const rows = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data = rows.slice(1);
    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    const output = [];

    // 送り主住所（sender.address）を2行に分割
    const senderAddrLines = splitByLength(sender.address, 25, 2);

    for (const r of data) {
      const out = Array(74).fill("");

      // ★ インポート元 CSV の列指定はテスト版通りそのまま
      const addrFull = r[12] || "";          // フル住所
      const postal   = cleanTelPostal(r[11] || "");
      const tel      = cleanTelPostal(r[14] || "");
      const name     = r[13] || "";
      const orderNo  = cleanOrderNumber(r[1] || "");

      // お届け先住所を3行まで 25文字分割
      const toAddrLines = splitByLength(addrFull, 25, 3);

      out[0]  = "0";          // A: コード取得区分
      out[2]  = tel;          // C: 電話番号
      out[3]  = postal;       // D: 郵便番号

      // E〜G: 住所1〜3 → 25文字分割
      out[4] = toAddrLines[0]; // 住所1
      out[5] = toAddrLines[1]; // 住所2
      out[6] = toAddrLines[2]; // 住所3

      out[7] = name;          // 名称1（宛名）
      out[25] = orderNo;       // 名称2（ご注文番号）

      // ご依頼主
      out[17] = sender.phone;              // R: ご依頼主電話
      out[18] = sender.postal;             // S: ご依頼主郵便
      out[19] = senderAddrLines[0];        // T: ご依頼主住所1（25文字）
      out[20] = senderAddrLines[1];        // U: ご依頼主住所2（25文字〜）
      out[21] = sender.name;               // V: ご依頼主名称1

      out[24] = "ブーケ加工品";           // Z: 品名1
      out[58] = todayStr;                  // BG: 出荷日

      output.push(out);
    }

    const csvTextOut =
      headers.join(",") + "\r\n" +
      output.map(r=>r.map(v=>`"${v ?? ""}"`).join(",")).join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }

  // ==========================================================
  // 🟦 ヤマト（B2クラウド／住所を25文字で2分割） 
  // ==========================================================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚 ヤマト変換開始（住所25文字分割）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data    = rows.slice(1);

    const res = await fetch("./js/newb2web_template1.xlsx");
    const wb = XLSX.read(await res.arrayBuffer(),{type:"array"});
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const header = XLSX.utils.sheet_to_json(sheet,{header:1})[0];

    function colLetter(i){
      let s=""; 
      while(i>=0){ s=String.fromCharCode(i%26+65)+s; i=Math.floor(i/26)-1; }
      return s;
    }
    function idx(key){
      return header.findIndex(h=>typeof h==="string"&&h.includes(key));
    }

    // ヘッダ内の対象列（テスト版のまま）
    const map = {
      order : idx("お客様管理番号"),
      type  : idx("送り状種類"),
      cool  : idx("クール区分"),
      ship1 : idx("出荷予定日"),
      ship2 : idx("出荷日"),
      tel   : idx("お届け先電話番号"),
      zip   : idx("お届け先郵便番号"),
      adr   : idx("お届け先住所"),
      apt   : idx("お届け先アパートマンション"),
      name  : idx("お届け先名"),
      honor : idx("敬称"),
      sTel  : idx("ご依頼主電話番号"),
      sZip  : idx("ご依頼主郵便番号"),
      sAdr  : idx("ご依頼主住所"),
      sApt  : idx("ご依頼主アパートマンション"),
      sName : idx("ご依頼主名"),
      item  : idx("品名１")
    };

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    let rowExcel = 2;

    function set(i,val){
      if(i < 0) return;
      sheet[colLetter(i)+rowExcel] = { v: val, t: "s" };
    }

    for(const r of data){
      // ★ インポート元 CSV の列指定はテスト版そのまま
      const order = cleanOrderNumber(r[1]  || ""); // ご注文番号
      const tel   = cleanTelPostal(r[14]   || ""); // 電話番号
      const zip   = cleanTelPostal(r[11]   || ""); // 郵便番号
      const adr   = r[12] || "";                  // フル住所
      const name  = r[13] || "";                  // 宛名

      // お届け先住所を 25文字 × 2 に分割
      const toAddrLines = splitByLength(adr, 25, 2);
      // ご依頼主住所も 25文字 × 2
      const senderAddrLines = splitByLength(sender.address, 25, 2);

      set(map.order, order);
      set(map.type, "0");
      set(map.cool, "0");
      set(map.ship1, todayStr);
      set(map.ship2, todayStr);

      set(map.tel, tel);
      set(map.zip, zip);

      // 住所＆アパートマンション
      set(map.adr, toAddrLines[0]); // 1行目
      set(map.apt, toAddrLines[1]); // 2行目（あれば）

      set(map.name, name);
      set(map.honor, "様");

      // ご依頼主
      set(map.sTel, sender.phone);
      set(map.sZip, sender.postal);
      set(map.sAdr, senderAddrLines[0]);
      set(map.sApt, senderAddrLines[1]);
      set(map.sName, sender.name);

      set(map.item, "ブーケ加工品");

      rowExcel++;
    }

    return wb;
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

      try {
        if (courier === "yamato") {
          mergedWorkbook = await convertToYamato(file, sender);
          convertedCSV   = null;
        } else if (courier === "japanpost") {
          convertedCSV   = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
        } else { // sagawa
          convertedCSV   = await convertToSagawa(file, sender);
          mergedWorkbook = null;
        }

        showMessage("✔ 変換完了（住所25文字分割版）", "success");
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
          courier === "japanpost" ? "yupack_import.csv" :
          courier === "sagawa"    ? "sagawa_import.csv" :
          "output.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = name;
        link.click();
        URL.revokeObjectURL(link.href);
      }
    });
  }
})();
