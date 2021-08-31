function AddToStock() {
  try {
    var ss = SpreadsheetApp.openById("mainSheet");
    var sStock = ss.getSheetByName("Stock");

    var infoRange = sStock.getRange("F6:F19");
    var info = infoRange.getValues();
    var logsRow = sStock.getRange("Stock_Logs_CurrRow").getValue();
    //Logger.log(logsRow);
    var date = new Date();
    var parsedInfo = [[date, info[2][0], info[0][0], info[1][0], info[3][0]]];

    Logger.log(info);
    Logger.log(parsedInfo);

    var range = sStock.getRange("Stock");
    var findType = range.createTextFinder(parsedInfo[0][2]).findNext();
    var stockRow = findType.getRow();

    var s = sStock.getRange("C" + stockRow);
    s.setValue(s.getValue() + parsedInfo[0][3]);

    Logger.log("M" + logsRow + ":Q" + logsRow);
    sStock.getRange("M" + logsRow + ":Q" + logsRow).setValues(parsedInfo);
    sStock.getRange("F6:F8").setValues([[""], [""], [""]]);
  } catch (e) {
    Logger.log(e);
  }
}
