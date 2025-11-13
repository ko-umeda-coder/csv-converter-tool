// ============================
// XLSXライブラリ読み込み待機
// ============================
const waitForXLSX = () => new Promise((resolve) => {
  const check = () => {
    if (window.XLSX) {
      console.log("✅ XLSX 読み込み完了");
      resolve();
    } else setTimeout(check, 50);
  };
  check();
});

// ============================
// メイン処理
// ============================
(async () => {
  await waitForXLSX();
  console.log("🔥【テスト版】main.js 起動（住所なし）");

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
      address: document.getElementById("senderAddress").value.trim(), // ←使わない
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
  // 🟩 住所ゼロ版：ゆうパック（72列固定）
  // ==========================================================
  async function convertToJapanPost(csvFile, sender) {
    console.log("📮【テスト】ゆうパック（住所なし）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l => l.split(","));
    const data    = rows.slice(1);

    const output = [];
    const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, "/");

    for (const r of data) {
      const name   = r[12] || "";
      const postal = cleanTelPostal(r[10] || "");
      const phone  = cleanTelPostal(r[13] || "");
      const orderNo = cleanOrderNumber(r[1] || "");

      const row = [];

      row.push("1","0","","","","","1");
      row.push(name, "様", "", postal);

      // ★住所関連 全て空欄 (12〜15列)
      row.push("", "", "", "");

      row.push(phone,"","","");
      row.push("","","");
      row.push(sender.name,"","",sender.postal);

      // 依頼主住所 全て空欄
      row.push("", "", "", "");

      row.push(sender.phone,"");
      row.push(orderNo,"");
      row.push("ブーケ加工品","","");
      row.push(todayStr,"","","","","");

      // 残り空欄
      while (row.length < 71) row.push("");
      row.push("0"); // 最後の列

      output.push(row);
    }

    const csvOut = output.map(r => r.map(v=>`"${v}"`).join(",")).join("\r\n");
    const sjis = Encoding.convert(Encoding.stringToCode(csvOut),"SJIS");
    return new Blob([new Uint8Array(sjis)],{type:"text/csv"});
  }


  // ==========================================================
  // 🟩 住所ゼロ版：佐川（74列固定）
  // ==========================================================
  async function convertToSagawa(csvFile, sender) {
    console.log("📦【テスト】佐川（住所なし）");

    const headers = [/* 74項目そのまま */];

    const csvText = await csvFile.text();
    const rows = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data = rows.slice(1);

    const todayStr = new Date().toISOString().slice(0, 10).replace(/-/g, "/");
    const output = [];

    for (const r of data) {
      const out = Array(74).fill("");

      out[0] = "0";
      out[2] = cleanTelPostal(r[13]||"");
      out[3] = cleanTelPostal(r[10]||"");

      // ★住所1/2/3 全て空欄（4,5,6）

      out[7] = r[12] || "";                 // 名称1
      out[8] = cleanOrderNumber(r[1] || ""); // 名称2（注文番号）

      out[17] = sender.phone;
      out[18] = sender.postal;

      // ご依頼主住所1/2 も空欄（19,20）
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
  // 🟩 住所ゼロ版：ヤマト（95列 Excel）
  // ==========================================================
  async function convertToYamato(csvFile, sender) {
    console.log("🚚【テスト】ヤマト（住所なし）");

    const csvText = await csvFile.text();
    const rows    = csvText.trim().split(/\r?\n/).map(l=>l.split(","));
    const data    = rows.slice(1);

    const res = await fetch("./js/newb2web_template1.xlsx");
    const wb  = XLSX.read(await res.arrayBuffer(), { type:"array" });

    const sheet = wb.Sheets[wb.SheetNames[0]];
    const header = XLSX.utils.sheet_to_json(sheet,{header:1})[0];

    function colLetter(i){let s="";while(i>=0){s=String.fromCharCode(i%26+65)+s;i=Math.floor(i/26)-1;}return s;}

    const idx = (kw)=>header.findIndex(h=>typeof h==="string" && h.includes(kw));

    const map = {
      order : idx("お客様管理番号"),
      type  : idx("送り状種類"),
      cool  : idx("クール区分"),
      ship1 : idx("出荷予定日"),
      ship2 : idx("出荷日"),
      deltel: idx("お届け先電話番号"),
      delzip: idx("お届け先郵便番号"),
      deladr: idx("お届け先住所"),
      delapt: idx("お届け先アパート"),
      delnam: idx("お届け先名"),
      honor : idx("敬称"),
      snttel: idx("ご依頼主電話番号"),
      sntzip: idx("ご依頼主郵便番号"),
      sntadr: idx("ご依頼主住所"),
      sntapt: idx("ご依頼主アパート"),
      sntnam: idx("ご依頼主名"),
      item  : idx("品名１"),
    };

    const todayStr = new Date().toISOString().slice(0,10).replace(/-/g,"/");

    let rowExcel = 2;
    function set(i,val){
      if(i<0)return;
      sheet[colLetter(i)+rowExcel]={v:val,t:"s"};
    }

    for(const r of data){
      const order = cleanOrderNumber(r[1]||"");
      const tel   = cleanTelPostal(r[13]||"");
      const zip   = cleanTelPostal(r[10]||"");
      const name  = r[12] || "";

      set(map.order, order);
      set(map.type, "0");
      set(map.cool, "0");
      set(map.ship1, todayStr);
      set(map.ship2, todayStr);

      set(map.deltel, tel);
      set(map.delzip, zip);

      // ★住所全削除
      set(map.deladr, "");
      set(map.delapt, "");

      set(map.delnam, name);
      set(map.honor, "様");

      set(map.snttel, sender.phone);
      set(map.sntzip, sender.postal);
      set(map.sntadr, "");
      set(map.sntapt, "");
      set(map.sntnam, sender.name);

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
      const file = fileInput.files[0];
      const courier = courierSelect.value;
      if (!file) return;

      const sender = getSenderInfo();
      showLoading(true);

      try {
        if (courier === "yamato") {
          mergedWorkbook = await convertToYamato(file, sender);
          convertedCSV = null;
        } else if (courier === "japanpost") {
          convertedCSV = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
        } else {
          convertedCSV = await convertToSagawa(file, sender);
          mergedWorkbook = null;
        }
        showMessage("✔ テスト出力完了", "success");
        downloadBtn.style.display = "block";
      } finally {
        showLoading(false);
      }
    });
  }

  function setupDownloadButton() {
    downloadBtn.addEventListener("click", () => {
      const courier = courierSelect.value;

      if (courier === "yamato" && mergedWorkbook) {
        XLSX.writeFile(mergedWorkbook, "yamato_test.xlsx");
        return;
      }

      if (convertedCSV) {
        const name = courier==="japanpost" ? "yupack_test.csv"
                  : courier==="sagawa"    ? "sagawa_test.csv"
                  : "output.csv";

        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = name;
        link.click();
      }
    });
  }
})();
