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
      if (name.startsWith("Review")) {
        applyReviewFormatting(sheet);
      }
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

function applyReviewFormatting(sheet) {
  const values = sheet.getDataRange().getValues();
  let headerRow = -1;
  let totalRow = -1;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "Category" && values[i][4] === "Status") {
      headerRow = i + 1;
    }
    if (headerRow > 0 && values[i][0] === "TOTAL") {
      totalRow = i + 1;
      break;
    }
  }

  if (headerRow < 0 || totalRow <= headerRow + 1) return;

  const firstCategoryRow = headerRow + 1;
  const categoryRowCount = totalRow - firstCategoryRow;
  const categoryRange = sheet.getRange(firstCategoryRow, 1, categoryRowCount, 6);
  const existingRules = sheet.getConditionalFormatRules().filter(rule => {
    return !rule.getRanges().some(range => range.getSheet().getName() === sheet.getName());
  });

  const redRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($E${firstCategoryRow}="Over",$E${firstCategoryRow}="Unbudgeted")`)
    .setBackground("#FCE4D6")
    .setRanges([categoryRange])
    .build();

  const greenRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=OR($E${firstCategoryRow}="Under",$E${firstCategoryRow}="OK")`)
    .setBackground("#E2F0D9")
    .setRanges([categoryRange])
    .build();

  sheet.setConditionalFormatRules(existingRules.concat([redRule, greenRule]));
}
