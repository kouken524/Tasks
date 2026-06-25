// 民宿大浜荘 — 予約台帳自動生成スクリプト
// 縦軸：部屋、横軸：日付

const ROOMS = ["広間", "2号", "3号", "5号", "6号", "7号", "8号"];

function createBookingCalendar() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "予約台帳を作成",
    "年月を入力してください（例：2025/07）",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const input = response.getResponseText().trim();
  const match = input.match(/^(\d{4})[\/\-](\d{1,2})$/);
  if (!match) {
    ui.alert("形式が正しくありません。「2025/07」のように入力してください。");
    return;
  }

  const year  = parseInt(match[1]);
  const month = parseInt(match[2]);
  const daysInMonth = new Date(year, month, 0).getDate();
  const sheetName = `${year}年${String(month).padStart(2,"0")}月 台帳`;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    const overwrite = ui.alert(
      `「${sheetName}」は既に存在します。上書きしますか？`,
      ui.ButtonSet.YES_NO
    );
    if (overwrite !== ui.Button.YES) return;
    sheet.clearContents();
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(sheetName);
  }

  // ヘッダー行を作る
  const weekdays = ["日","月","火","水","木","金","土"];
  const headerRow = ["部屋"];
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    headerRow.push(`${d}(${weekdays[dow]})`);
  }
  headerRow.push("備考");

  // 部屋行を作る
  const rows = [headerRow];
  ROOMS.forEach(room => {
    const row = [room];
    for (let d = 1; d <= daysInMonth; d++) row.push("");
    row.push("");
    rows.push(row);
  });

  // シートに書き込む
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  // 書式設定
  const totalCols = headerRow.length;
  const totalRows = rows.length;

  // ヘッダー行
  const headerRange = sheet.getRange(1, 1, 1, totalCols);
  headerRange.setFontWeight("bold").setBackground("#1565C0").setFontColor("#FFFFFF");

  // 部屋列
  sheet.getRange(2, 1, ROOMS.length, 1)
       .setFontWeight("bold")
       .setBackground("#E3F2FD")
       .setHorizontalAlignment("center");

  // 土日に色をつける
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    if (dow === 0 || dow === 6) {
      const col = d + 1;
      const bg = dow === 0 ? "#FFEBEE" : "#FFF3E0";
      sheet.getRange(1, col, totalRows, 1).setBackground(bg);
    }
  }

  // 全体の罫線・中央揃え
  sheet.getRange(1, 1, totalRows, totalCols)
       .setBorder(true, true, true, true, true, true)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");

  // 列幅を調整
  sheet.setColumnWidth(1, 70);
  for (let c = 2; c <= totalCols - 1; c++) sheet.setColumnWidth(c, 52);
  sheet.setColumnWidth(totalCols, 100);

  // 行の高さ
  for (let r = 2; r <= totalRows; r++) sheet.setRowHeight(r, 40);

  sheet.setFrozenColumns(1);
  sheet.setFrozenRows(1);

  ui.alert(`「${sheetName}」を作成しました！\n予約はセルに直接入力してください。`);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🏠 大浜荘 予約管理")
    .addItem("月別予約台帳を作成", "createBookingCalendar")
    .addToUi();
}
