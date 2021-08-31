var vars = {
  sheet: {
    id: "mainSheet",
    name: "Company Money",
  },
  columns: {
    dataStart: 9,
    dataEnd: 4,
  },
  ranges: {
    data: "CompanyMoney_Data",
    money: "CompanyMoney_Money",
    nextRow: "CompanyMoney_NextRow",
  },
};

function depositMoney() {
  try {
    var sheet = SpreadsheetApp.openById(vars["sheet"]["id"]).getSheetByName(vars["sheet"]["name"]);
    var currentMoneyRange = sheet.getRange(vars["ranges"]["money"]);
    var dataRange = sheet.getRange(vars["ranges"]["data"]);

    var nextRow = sheet.getRange(vars["ranges"]["nextRow"]).getValue();
    var data = dataRange.getValues();
    var currentMoney = currentMoneyRange.getValue();

    if (data[0] != "" && data[1] != "" && data[2] != "") {
      var amount = data[0][0];
      var worker = data[1][0];
      var reason = data[2][0];

      currentMoneyRange.setValue(currentMoney + amount);

      var date = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy HH:mm");
      var logToPaste = [[date, worker, amount, reason]];

      sheet.getRange(nextRow, 11).setFontColor("green");
      sheet.getRange(nextRow, vars["columns"]["dataStart"], 1, vars["columns"]["dataEnd"]).setValues(logToPaste);

      dataRange.setValues([[""], [""], [""], [""]]);
    }
  } catch (e) {
    Logger.log(e);
  }
}

function withdrawMoney() {
  try {
    var sheet = SpreadsheetApp.openById(vars["sheet"]["id"]).getSheetByName(vars["sheet"]["name"]);
    var currentMoneyRange = sheet.getRange(vars["ranges"]["money"]);
    var dataRange = sheet.getRange(vars["ranges"]["data"]);

    var nextRow = sheet.getRange(vars["ranges"]["nextRow"]).getValue();
    var data = dataRange.getValues();
    var currentMoney = currentMoneyRange.getValue();

    if (data[0] != "" && data[1] != "" && data[2] != "") {
      var amount = data[0][0];
      var worker = data[1][0];
      var reason = data[2][0];

      currentMoneyRange.setValue(currentMoney - amount);

      var date = Utilities.formatDate(new Date(), "GMT", "dd/MM/yyyy HH:mm");
      amount = "-" + amount;
      var logToPaste = [[date, worker, amount, reason]];

      sheet.getRange(nextRow, 11).setFontColor("red");
      sheet.getRange(nextRow, vars["columns"]["dataStart"], 1, vars["columns"]["dataEnd"]).setValues(logToPaste);

      dataRange.setValues([[""], [""], [""], [""]]);
    }
  } catch (e) {
    Logger.log(e);
  }
}
