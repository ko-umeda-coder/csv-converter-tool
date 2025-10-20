console.log("main.js ロード完了");

// ユーザー画面簡易処理
document.getElementById('convertBtn')?.addEventListener('click', ()=>{
  const file = document.getElementById('csvFile').files[0];
  const courier = document.getElementById('courierSelect').value;
  if(!file){ alert("CSVを選択してください"); return; }
  alert(`CSVを読み込み、配送会社:${courier} で変換（最小構成）`);
  document.getElementById('preview').textContent = `CSVプレビュー（最小構成）: ${file.name}`;
});

// 管理者画面簡易処理
document.getElementById('updateFormatBtn')?.addEventListener('click', ()=>{
  const file = document.getElementById('adminFile').files[0];
  const courier = document.getElementById('adminCourierSelect').value;
  if(!file){ alert("CSV/Excelを選択してください"); return; }
  alert(`管理者モード：${courier} のフォーマットを更新（最小構成）`);
  document.getElementById('adminPreview').textContent = `読み込んだファイル: ${file.name}`;
});

