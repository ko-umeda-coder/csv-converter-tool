// ============================
// main.js 完全安定版（列ズレ対策済）
// ============================
document.addEventListener("DOMContentLoaded", () => {
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

  if (typeof Papa === "undefined" || typeof Encoding === "undefined") {
    console.error("❌ 必要なライブラリが読み込まれていません。");
    showMessage("初期化エラー: 必要なライブラリが見つかりません。", "error");
    return;
  }

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
      { value: "yamato", text: "ヤマト運輸" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
    ];
    const initialOption =
      '<option value="" disabled selected>--- 選択してください ---</option>';
    courierSelect.innerHTML =
      initialOption +
      options.map((o) => `<option value="${o.value}">${o.text}</option>`).join("");
  }

  // ============================
  // ファイル選択＆ドラッグドロップ
  // ============================
  function setupFileInput() {
    const updateFileState = (file) => {
      if (file) {
        console.log(`✅ ${file.name} が選択されました`);
        fileName.textContent = file.name;
        fileWrapper.classList.add("has-file");
        convertBtn.disabled = false;
      } else {
        console.warn("⚠ ファイルが選択されていません");
        fileName.textContent = "";
        fileWrapper.classList.remove("has-file");
        convertBtn.disabled = true;
      }
    };

    fileInput.addEventListener("change", (e) => {
      updateFileState(e.target.files?.[0]);
    });

    fileWrapper.addEventListener("dragover", (e) => {
      e.preventDefault();
      fileWrapper.classList.add("drag-over");
    });
    fileWrapper.addEventListener("dragleave", (e) => {
      e.preventDefault();
      fileWrapper.classList.remove("drag-over");
    });
    fileWrapper.addEventListener("drop", (e) => {
      e.preventDefault();
      fileWrapper.classList.remove("drag-over");
      const file = e.dataTransfer.files[0];
      if (file && file.name.endsWith(".csv")) {
        fileInput.files = e.dataTransfer.files;
        updateFileState(file);
      } else if (file) {
        showMessage("対応していないファイル形式です。CSVファイルを選択してください。", "error");
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
      overlay.innerHTML = `
        <div class="loading-content">
          <div class="spinner"></div>
          <div class="loading-text">処理中...</div>
        </div>`;
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

  function validateSenderInfo(sender) {
    if (!sender.name || !sender.postal || !sender.address || !sender.phone) {
      return "送り主情報の入力欄は全て必須です。ご確認ください。";
    }
    return null;
  }

  // ============================
  // クレンジング関数群
  // ============================
  function cleanTelPostal(v) {
    if (!v) return "";
    return String(v)
      .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xFEE0))
      .replace(/[ー－―]/g, "-")
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

  function cleanText(v) {
    if (!v) return "";
    return String(v)
      .replace(/[　\s]+/g, " ")
      .replace(/["']/g, "")
      .trim();
  }

  // ============================
  // CSV解析 (PapaParse)
  // ============================
  async function parseCSV(csvFile) {
    const text = await csvFile.text();
    const parsed = Papa.parse(text, {
      header: false,
      skipEmptyLines: true,
      encoding: "Shift-JIS",
      quoteChar: '"',
      delimiter: ",",
    });

    if (parsed.errors.length > 0) {
      console.error("CSV解析エラー:", parsed.errors);
      throw new Error(`CSVの解析中にエラーが発生しました: ${parsed.errors[0].message}`);
    }

    const dataRows = parsed.data.slice(1);
    if (dataRows.length === 0) {
      throw new Error("CSVファイルにデータ行がありません。");
    }

    return dataRows;
  }

  // ============================
  // ヤマト運輸変換処理
  // ============================
  async function mergeToYamatoTemplate(csvFile, templateUrl, sender) {
    const dataRows = await parseCSV(csvFile);
    const res = await fetch(templateUrl);
    if (!res.ok) throw new Error(`テンプレートファイルが見つかりません: ${templateUrl}`);

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets["外部データ取り込み基本レイアウト"];
    if (!sheet) throw new Error("テンプレートに指定シートが見つかりません。");

    let rowExcel = 2;
    for (const r of dataRows) {
      const orderNumber = cleanOrderNumber(r[1]);
      const postal = cleanTelPostal(r[10]);
      const addressFull = cleanText(r[11] || "");
      const name = cleanText(r[12] || "");
      const phone = cleanTelPostal(r[13]);
      sheet[`A${rowExcel}`] = { v: orderNumber, t: "s" };
      sheet[`K${rowExcel}`] = { v: postal, t: "s" };
      sheet[`L${rowExcel}`] = { v: addressFull, t: "s" };
      sheet[`P${rowExcel}`] = { v: name, t: "s" };
      sheet[`I${rowExcel}`] = { v: phone, t: "s" };
      sheet[`AB${rowExcel}`] = { v: "ブーケ加工品", t: "s" };
      rowExcel++;
    }
    return wb;
  }

  // ============================
  // ゆうプリR変換処理（列ズレ完全対策版）
  // ============================
  async function convertToJapanPost(csvFile, sender) {
    const dataRows = await parseCSV(csvFile);
    const output = [];

    for (const r of dataRows) {
      const rowOut = new Array(73).fill("");

      // 固定値
      rowOut[0] = "1"; // A列
      rowOut[1] = "0"; // B列
      rowOut[6] = "1"; // G列（敬称コード）
      rowOut[67] = "0"; // BM列
      rowOut[72] = "0"; // BT列

      const orderNo = cleanOrderNumber(r[1] || "");
      const name = cleanText(r[12] || "");
      const postal = cleanTelPostal(r[10] || "");
      const address = cleanText(r[11] || "");
      const phone = cleanTelPostal(r[13] || "");

      // 宛先情報
      rowOut[7] = name;
      rowOut[10] = postal;
      if (address.length > 25) {
        rowOut[11] = address.slice(0, 25);
        rowOut[12] = address.slice(25);
      } else {
        rowOut[11] = address;
      }
      rowOut[15] = phone;

      // 送り主情報
      const senderAddress = cleanText(sender.address || "");
      rowOut[22] = cleanText(sender.name || "");
      rowOut[25] = cleanTelPostal(sender.postal || "");
      if (senderAddress.length > 25) {
        rowOut[26] = senderAddress.slice(0, 25);
        rowOut[27] = senderAddress.slice(25);
      } else {
        rowOut[26] = senderAddress;
      }
      rowOut[30] = cleanTelPostal(sender.phone || "");

      // 注文番号・品名
      rowOut[32] = orderNo;
      rowOut[34] = "ブーケ加工品";

      output.push(rowOut);
    }

    // === 列数を固定してCSV生成 ===
    const normalizedOutput = output.map((row) => {
      const fixed = [...row];
      while (fixed.length < 73) fixed.push("");
      return fixed.slice(0, 73);
    });

    const csvText = normalizedOutput
      .map((row) =>
        row.map((v) => `"${String(v).replace(/"/g, '""')}"`).join(",")
      )
      .join("\r\n");

    const sjis = Encoding.convert(Encoding.stringToCode(csvText), "SJIS");
    return new Blob([new Uint8Array(sjis)], { type: "text/csv;charset=Shift_JIS" });
  }

  // ============================
  // 変換ボタン
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file = fileInput.files?.[0];
      const courier = courierSelect.value;
      const sender = getSenderInfo();

      if (!file || !courier) {
        showMessage("ファイルを選択し、宅配会社を選んでください。", "error");
        return;
      }
      const validationError = validateSenderInfo(sender);
      if (validationError) {
        showMessage(validationError, "error");
        return;
      }

      showLoading(true);
      showMessage("変換処理中...", "info");
      downloadBtn.disabled = true;

      try {
        if (courier === "japanpost") {
          convertedCSV = await convertToJapanPost(file, sender);
          mergedWorkbook = null;
          showMessage("✅ ゆうプリR変換完了 (ダウンロード可能)", "success");
        } else {
          mergedWorkbook = await mergeToYamatoTemplate(file, "./js/newb2web_template1.xlsx", sender);
          convertedCSV = null;
          showMessage("✅ ヤマト変換完了 (ダウンロード可能)", "success");
        }

        downloadBtn.style.display = "block";
        downloadBtn.disabled = false;
        downloadBtn.className = "btn btn-primary";
      } catch (err) {
        console.error("変換処理エラー:", err);
        showMessage(`変換中にエラーが発生しました: ${err.message}`, "error");
      } finally {
        showLoading(false);
      }
    });
  }

  // ============================
  // ダウンロード処理
  // ============================
  function setupDownloadButton() {
    downloadBtn.disabled = true;
    downloadBtn.addEventListener("click", () => {
      if (mergedWorkbook) {
        XLSX.writeFile(mergedWorkbook, "yamato_b2_import.xlsx");
        showMessage("ヤマト用ファイルをダウンロードしました。", "info");
      } else if (convertedCSV) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = "yupack_import.csv";
        link.click();
        URL.revokeObjectURL(link.href);
        showMessage("ゆうプリR用ファイルをダウンロードしました。", "info");
      } else {
        showMessage("変換データがありません。先に変換を実行してください。", "error");
      }
    });
  }
});
