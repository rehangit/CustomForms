function CustomForms_doPost(e: GoogleAppsScript.Events.DoPost) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = ss.getSheetByName("logs");
  if (logs) logs.appendRow([new Date(), "POST", JSON.stringify(e)]);

  const pathInfo: string = e.pathInfo;
  if (pathInfo && pathInfo.length) {
    const target = ss.getSheetByName(pathInfo) || ss.insertSheet(pathInfo);

    const obj = { Timestamp: new Date(), ...e.parameter };
    const keys = Object.keys(obj);
    const firstRow = target
      .getDataRange()
      .getValues()[0]
      .map(v => v.toString().trim())
      .filter(Boolean);
    const firstRowLength = (firstRow && firstRow.filter(Boolean).length) || 0;
    const lastRow = firstRow.map(() => "");

    const firstRowExtra = [];
    keys.forEach((k: string) => {
      const v = obj[k];
      const ik = firstRow.indexOf(k);
      if (ik >= 0) lastRow[ik] = v;
      else {
        firstRowExtra.push(k);
        lastRow.push(v);
      }
    });

    if (firstRowExtra.length)
      target.getRange(1, firstRowLength + 1, 1, firstRowExtra.length).setValues([firstRowExtra]);
    target.appendRow(lastRow);
  }

  return ContentService.createTextOutput(JSON.stringify(e.parameter));
}

function CustomForms_doGet(e: GoogleAppsScript.Events.DoGet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = ss.getSheetByName("logs");
  if (logs) logs.appendRow([new Date(), "GET", JSON.stringify(e)]);

  const keyValues = ss.getSheetByName(e.parameter["source"]);

  const data = JSON.stringify(
    (keyValues &&
      keyValues
        .getDataRange()
        .getValues()
        .filter(function(row) {
          return row[0].length > 0;
        })) || [[]]
  );
  const callback = e.parameter["callback"];
  return callback
    ? ContentService.createTextOutput(callback + "(" + data + ")").setMimeType(
        ContentService.MimeType.JAVASCRIPT
      )
    : ContentService.createTextOutput(data).setMimeType(ContentService.MimeType.JSON);
}
