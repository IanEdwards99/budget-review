function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (!payload.spreadsheetId) throw new Error("Missing spreadsheetId");
    if (!payload.sheets || !Array.isArray(payload.sheets)) throw new Error("Missing sheets array");

    const ss = SpreadsheetApp.openById(payload.spreadsheetId);
    payload.sheets.forEach(sheetPayload => {
      const name = sheetPayload.name;
      const values = sheetPayload.values || [];
      let sheet = ss.getSheetByName(name);
      if (!sheet) sheet = ss.insertSheet(name);
      sheet.clear({ contentsOnly: false });

      const rowCount = Math.max(values.length, 1);
      const colCount = Math.max(...values.map(row => row.length), 1);
      if (sheet.getMaxRows() < rowCount) {
        sheet.insertRowsAfter(sheet.getMaxRows(), rowCount - sheet.getMaxRows());
      }
      if (sheet.getMaxColumns() < colCount) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), colCount - sheet.getMaxColumns());
      }

      const rectangular = values.map(row => {
        const next = row.slice();
        while (next.length < colCount) next.push("");
        return next;
      });

      if (rectangular.length > 0 && colCount > 0) {
        sheet.getRange(1, 1, rectangular.length, colCount).setValues(rectangular);
      }

      sheet.setFrozenRows(name.startsWith("Review") ? 11 : 1);
      sheet.autoResizeColumns(1, Math.min(colCount, 12));
    });

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: "Google Sheet updated" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
