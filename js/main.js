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
  console.log("🔥【テスト版】main.js（住所1列固定）起動");

  const fileInput     = document.getElementById("csvFile");
  const fileWrapper   = document.getElementById("fileWrapper");
  const fileName      = document.getElementById("fileName");
  const convertBtn    = document.getElementById("convertBtn");
  const downloadBtn   = document.getElementById("downloadBtn");
  const messageBox    = document.getElementById("message");
  const courierSelect = document.getElementById("courierSelect");

  let mergedWorkbook = null;
  let convertedCSV   = null;

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
      address: document.getElementById("senderAddress").value.trim(), // ← 1列としてそのまま使用
      phone:   cleanTelPostal(document.getElementById("senderPhone").value.trim()),
    };
  }

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

  function cleanTelPostal(v) {
    if (!v) return "";
    return String(v).replace(/[^0-9\-]/g, "");
  }
  function cleanOrderNumber(v) {
    if (!v) return "";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "");
  }

  // ==========================================================
  // 🟥 ゆうパック（住所1列・72列固定）
  // ==========================================================
  async function convertToJapanPost(csvFile, sender) {
    console.log("📮【テスト】ゆうパック開始（住所1列）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data    = rows.slice(1);

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");
    const output = [];

    for (const r of data) {
      const name = r[11] || "";
      const postal = cleanTelPostal(r[9] || "");
      const addrFull = r[12] || "";   // ← フル住所1列
      const phone = cleanTelPostal(r[12] || "");
      const orderNo = cleanOrderNumber(r[1] || "");

      const row = [];

      row.push("1","0","","","","","1"); // 1〜7
      row.push(postal);      // 8
      row.push("様");      // 9
      row.push("");        // 10
      row.push(name);    // 11

      // 12〜15（住所） → addrFull のみを入れて残り空白
      row.push(phone);  // 12
      row.push("");        // 13
      row.push("");        // 14
      row.push("");        // 15

      row.push(phone); row.push(""); row.push(""); row.push(""); // 16〜19

      // ...略（依頼主情報）
      row.push(sender.name,"","",sender.postal);    // 23〜26
      row.push(sender.address);                     // 27（住所1列）
      row.push("");                                 // 28
      row.push("");                                 // 29
      row.push("");                                 // 30

      row.push(sender.phone,"",orderNo,"");         // 31〜34

      row.push("ブーケ加工品","","");               // 35〜37

      row.push(todayStr);                           // 38
      row.push("","","","","");                     // 39〜43

      // 残り埋める
      while (row.length < 71) row.push("");
      row.push("0"); // 72列目（配達完了通知（依頼主））

      output.push(row);
    }

    const csvOut = output.map(r => r.map(v=>`"${v}"`).join(",")).join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }

  // ==========================================================
  // 🟥 佐川（住所1列・74列固定）
  // ==========================================================
  async function convertToSagawa(csvFile, sender) {
    console.log("📦【テスト】佐川開始（住所1列）");

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

    for (const r of data) {
      const out = Array(74).fill("");

      const addrFull = r[13] || "";
      const postal   = cleanTelPostal(r[12] || "");

      out[0]  = "0";
      out[2]  = cleanTelPostal(r[15]||"");
      out[3]  = postal;

      // 住所1のみにセット（住所2,3 は空欄）
      out[4] = addrFull;
      out[5] = "";
      out[6] = "";

      out[7] = r[14] || "";
      out[8] = cleanOrderNumber(r[1] || "");

      out[17] = sender.phone;
      out[18] = sender.postal;

      // ご依頼主住所1 のみに sender.address
      out[19] = sender.address;
      out[20] = "";

      out[21] = sender.name;

      out[25] = "ブーケ加工品";
      out[58] = todayStr;

      output.push(out);
    }

    const csvTextOut =
      headers.join(",") + "\r\n" +
      output.map(r=>r.map(v=>`"${v}"`).join(",")).join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvTextOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }

  // ==========================================================
  // 🟥 ヤマト（住所1列・Excel）
  // ==========================================================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚【テスト】ヤマト開始（住所1列）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data    = rows.slice(1);

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
      type  : idx("送り状種類"),
      cool  : idx("クール区分"),
      ship1 : idx("出荷予定日"),
      ship2 : idx("出荷日"),
      tel   : idx("お届け先電話番号"),
      zip   : idx("お届け先郵便番号"),
      adr   : idx("お届け先住所"),
      apt   : idx("お届け先アパート"),
      name  : idx("お届け先名"),
      honor : idx("敬称"),
      sTel  : idx("ご依頼主電話番号"),
      sZip  : idx("ご依頼主郵便番号"),
      sAdr  : idx("ご依頼主住所"),
      sApt  : idx("ご依頼主アパート"),
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
      const order = cleanOrderNumber(r[1]||"");
      const tel   = cleanTelPostal(r[14]||"");
      const zip   = cleanTelPostal(r[11]||"");
      const name  = r[13]||"";
      const adr   = r[12]||"";  // ★住所1列

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
          convertedCSV = null;
        } else if (courier === "japanpost") {
          convertedCSV   = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
        } else {
          convertedCSV   = await convertToSagawa(file, sender);
          mergedWorkbook = null;
        }

        showMessage("✔ テスト出力完了（住所1列版）", "success");
        downloadBtn.style.display = "block";

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
        XLSX.writeFile(mergedWorkbook, "yamato_test.xlsx");
        return;
      }

      if (convertedCSV) {
        const name =
          courier === "japanpost" ? "yupack_test.csv" :
          courier === "sagawa"    ? "sagawa_test.csv" :
          "output.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = name;
        link.click();
      }
    });
  }

})();
