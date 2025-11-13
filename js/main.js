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
  // 初期設定
  // ============================
  setupCourierOptions();
  setupFileInput();
  setupConvertButton();
  setupDownloadButton();

  // 宅配会社リスト
  function setupCourierOptions() {
    const options = [
      { value: "yamato", text: "ヤマト運輸（B2クラウド）" },
      { value: "japanpost", text: "日本郵政（ゆうプリR）" },
      { value: "sagawa", text: "佐川急便（e飛伝Ⅱ）" }
    ];
    courierSelect.innerHTML = options.map(o => `<option value="${o.value}">${o.text}</option>`).join("");
  }

  // ファイル選択
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

  // メッセージ表示
  function showMessage(text, type = "info") {
    messageBox.style.display = "block";
    messageBox.textContent = text;
    messageBox.className = "message " + type;
  }

  // ローディング表示
  function showLoading(show) {
    let overlay = document.getElementById("loading");
    if (!overlay) {
      overlay = document.createElement("div");
      overlay.id = "loading";
      overlay.className = "loading-overlay";
      overlay.innerHTML =
        `<div class="loading-content"><div class="spinner"></div><div class="loading-text">変換中...</div></div>`;
      document.body.appendChild(overlay);
    }
    overlay.style.display = show ? "flex" : "none";
  }

  // 送り主情報取得
  function getSenderInfo() {
    return {
      name: document.getElementById("senderName").value.trim(),
      postal: document.getElementById("senderPostal").value.trim(),
      address: document.getElementById("senderAddress").value.trim(),
      phone: document.getElementById("senderPhone").value.trim(),
    };
  }

  // クレンジング
  function cleanTelPostal(v) {
    if (!v) return "0";
    return String(v).replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "").trim();
  }
  function cleanOrderNumber(v) {
    if (!v) return "0";
    return String(v).replace(/^(FAX|EC)/i, "").replace(/[★\[\]\s]/g, "").trim();
  }

  // 住所分割
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
    const rest = address.replace(pref, "");
    const [city, ...restParts] = rest.split(/(?<=市|区|町|村)/);
    const restFull = restParts.join("");
    const [rest1, building] = restFull.split(/[\s　]+/, 2);
    return { pref, city, rest: rest1 || "", building: building || "" };
  }

// ============================
// 佐川急便（e飛伝Ⅱ） ヘッダ名マッピング対応版
// ============================

async function convertToSagawa(csvFile, sender) {
  console.log("🚚 佐川変換処理（ヘッダ名マッピング方式）開始");

  // ① 入力 CSV（発送伝票対象一覧）読み込み
  const text = await csvFile.text();
  const rows = text.trim().split(/\r?\n/).map(line => line.split(","));

  const inputHeaders = rows[0]; // ← ヘッダ行
  const dataRows = rows.slice(1);

  // ② 佐川テンプレート（取り込み用CSV）読み込み（ヘッダ）
  const res = await fetch("./js/okurijo_torikomi_format.csv");
  const tmplText = await res.text();
  const tmplRows = tmplText.trim().split(/\r?\n/).map(line => line.split(","));
  const outputHeaders = tmplRows[0]; // ← 正しい版のヘッダ行
  const totalCols = outputHeaders.length;

  console.log("入力ヘッダ：", inputHeaders);
  console.log("出力ヘッダ：", outputHeaders);

  // ③ 入力CSVのヘッダ → index 変換
  const inputIndex = {};
  inputHeaders.forEach((h, idx) => (inputIndex[h.trim()] = idx));

  // ④ マッピングルール（A〜BV の仕様をヘッダ名で定義）
  const mapping = {
    "お届け先コード取得区分": { value: "0" },
    "お届け先コード": {},
    "お届け先電話番号": { from: "電話番号（半角英数）", clean: "tel" },
    "お届け先郵便番号": { from: "郵便番号（半角英数）", clean: "postal" },
    "お届け先住所１": { from: "住所（都道府県・建物名含む）", split: "prefCity" },
    "お届け先住所２": { from: "住所（都道府県・建物名含む）", split: "rest" },
    "お届け先住所３": { from: "住所（都道府県・建物名含む）", split: "building" },
    "お届け先名称１": { from: "お届け先の宛名" },
    "お届け先名称２": { from: "ご注文番号", clean: "order" },

    "お客様管理番号": {},
    "お客様コード": {},
    "部署ご担当者コード取得区分": {},
    "部署ご担当者コード": {},
    "部署ご担当者名称": {},
    "荷送人電話番号": {},

    "ご依頼主コード取得区分": {},
    "ご依頼主コード": {},
    "ご依頼主電話番号": { fromSender: "phone", clean: "tel" },
    "ご依頼主郵便番号": { fromSender: "postal", clean: "postal" },
    "ご依頼主住所１": { fromSender: "address", split: "prefCity" },
    "ご依頼主住所２": { fromSender: "address", split: "rest" },
    "ご依頼主名称１": { fromSender: "name" },
    "ご依頼主名称２": {},

    "荷姿": {},
    "品名１": { value: "ブーケ加工品" },
    "品名２": {},
    "品名３": {},
    "品名４": {},
    "品名５": {},

    // 荷札関係
    "荷札荷姿": {},
    "荷札品名１": {},
    "荷札品名２": {},
    "荷札品名３": {},
    "荷札品名４": {},
    "荷札品名５": {},
    "荷札品名６": {},
    "荷札品名７": {},
    "荷札品名８": {},
    "荷札品名９": {},
    "荷札品名10": {},
    "荷札品名11": {},

    "出荷個数": {},
    "スピード指定": {},
    "クール便指定": {},
    "配達日": {},

    "配達指定時間帯": {},
    "配達指定時間（時分）": {},
    "代引金額": {},
    "消費税": {},
    "決済種別": {},
    "保険金額": {},

    "指定シール1": {},
    "指定シール2": {},
    "指定シール3": {},
    "営業所受取": {},
    "SRC区分": {},
    "営業所受取営業所コード": {},
    "元着区分": {},
    "メールアドレス": {},
    "ご不在時連絡先": {},

    "出荷日": { value: "TODAY" },
    "お問い合せ送り状No.": {},
    "出荷場印字区分": {},
    "集約解除指定": {},

    "編集01": {},
    "編集02": {},
    "編集03": {},
    "編集04": {},
    "編集05": {},
    "編集06": {},
    "編集07": {},
    "編集08": {},
    "編集09": {},
    "編集10": {}
  };

  // ⑤ 住所分割関数
  function splitAddr(text) {
    if (!text) return { prefCity: "", rest: "", building: "" };
    const prefList = ["東京都","北海道","京都府","大阪府","神奈川県","千葉県","埼玉県",
      "愛知県","兵庫県","福岡県","静岡県","茨城県","広島県","宮城県","新潟県",
      "長野県","岐阜県","群馬県","栃木県","岡山県","熊本県","滋賀県","三重県",
      "鹿児島県","山口県","愛媛県","奈良県","青森県","沖縄県","石川県","香川県",
      "大分県","岩手県","山形県","富山県","福島県","佐賀県","秋田県","山梨県","福井県","和歌山県","徳島県","高知県"];

    const pref = prefList.find(p => text.startsWith(p)) || "";
    let rest = text.replace(pref, "");
    const cityMatch = rest.match(/^(.*?[市区町村])/);
    const city = cityMatch ? cityMatch[1] : "";
    rest = rest.replace(city, "");

    // 丁番地と建物名をゆるく分割
    const bldgMatch = rest.match(/(.*?)(ビル|マンション|ハイツ|荘|号室|階|F).*/);
    const restOnly = bldgMatch ? bldgMatch[1].trim() : rest.trim();
    const building = bldgMatch ? rest.replace(restOnly, "").trim() : "";

    return {
      prefCity: pref + city,
      rest: restOnly,
      building: building
    };
  }

  // ⑥ クレンジング
  function clean(val, type) {
    if (!val) return "";
    let v = String(val).trim();

    if (type === "tel" || type === "postal") {
      v = v.replace(/^="?/, "").replace(/"$/, "").replace(/[^0-9\-]/g, "");
    }
    if (type === "order") {
      v = v.replace(/^(FAX|EC)/, "").replace(/[★\[\]\s]/g, "");
    }
    return v;
  }

  // ⑦ 行変換
  const output = [];

  for (const r of dataRows) {
    const out = Array(totalCols).fill("");

    outputHeaders.forEach((header, colIndex) => {
      const rule = mapping[header];
      if (!rule) return;

      let value = "";

      // 固定値
      if (rule.value === "TODAY") {
        const d = new Date();
        value = `${d.getFullYear()}/${String(d.getMonth()+1).padStart(2,"0")}/${String(d.getDate()).padStart(2,"0")}`;
      } else if (rule.value !== undefined) {
        value = rule.value;
      }

      // 入力CSVから取得
      if (rule.from) {
        const idx = inputIndex[rule.from];
        if (idx !== undefined) {
          value = r[idx];
        }
      }

      // 送り主情報
      if (rule.fromSender) {
        value = sender[rule.fromSender] || "";
      }

      // クレンジング
      if (rule.clean) {
        value = clean(value, rule.clean);
      }

      // 住所分割
      if (rule.split) {
        const source = rule.fromSender ? sender.address : (r[inputIndex["住所（都道府県・建物名含む）"]] || "");
        const addr = splitAddr(source);
        value = addr[rule.split] || "";
      }

      out[colIndex] = value;
    });

    output.push(out);
  }

  // ⑧ CSV生成（SJIS）
  const csvOut =
    [outputHeaders.join(",")]
      .concat(output.map(r => r.map(v => `"${v}"`).join(",")))
      .join("\r\n");

  const sjisArray = Encoding.convert(Encoding.stringToCode(csvOut), "SJIS");
  return new Blob([new Uint8Array(sjisArray)], { type: "text/csv" });
}


  // ============================
  // ボタン処理
  // ============================
  function setupConvertButton() {
    convertBtn.addEventListener("click", async () => {
      const file = fileInput.files[0];
      if (!file) return;
      const courier = courierSelect.value;
      showLoading(true);
      try {
        const sender = getSenderInfo();

        if (courier === "sagawa") {
          convertedCSV = await convertToSagawa(file, sender);
          mergedWorkbook = null;
          showMessage("✅ 佐川急便（e飛伝Ⅱ）変換完了", "success");
        } else {
          showMessage("❌ 今は佐川のみ検証対象です", "error");
          return;
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
      if (convertedCSV) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(convertedCSV);
        link.download = "sagawa_import.csv";
        link.click();
        URL.revokeObjectURL(link.href);
      } else {
        alert("変換データがありません。");
      }
    });
  }
})();
