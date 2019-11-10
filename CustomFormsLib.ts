function log(obj: any, func?: string) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logs = ss.getSheetByName("logs");
  if (logs) {
    const str = obj instanceof String ? obj : JSON.stringify(obj);
    logs.appendRow([new Date(), "CustomForms:" + func, str]);
  }
}

function CustomForms_doPost(e: GoogleAppsScript.Events.DoPost) {
  log(e, "doPost");
  const returnValue = [];

  const destSheetName: string =
    e.pathInfo && e.pathInfo.length
      ? e.pathInfo
      : e.queryString.indexOf("_target=") >= 0
      ? e.parameter["_target"]
      : "";

  // log({ destSheetName }, "doPost");

  if (destSheetName && destSheetName.length) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const target = ss.getSheetByName(destSheetName) || ss.insertSheet(destSheetName);

    const contents = e.parameter;
    // log({ contents }, "doPost");

    const obj = { Timestamp: new Date(), ...contents };
    const keys = Object.keys(obj);
    const firstRow = target
      .getDataRange()
      .getValues()[0]
      .map(v => v.toString().trim())
      .filter(Boolean);
    const firstRowLength = (firstRow && firstRow.length) || 0;
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

    const values = target.getDataRange().getValues();
    returnValue.push(...values.slice(-4));
    returnValue[0] = values[0];
  }

  return ContentService.createTextOutput(JSON.stringify(returnValue));
}

function CustomForms_doGet(e: GoogleAppsScript.Events.DoGet) {
  log(e, "doGet");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const keyValues = ss.getSheetByName(e.parameter["_source"]);

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
