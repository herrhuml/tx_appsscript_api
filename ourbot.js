var sheet = null;
var data = null;

var fu = {
  staffList: {
    sheet: {
      id: "staffList",
      name: "Staff List",
      nameInactive: "Inactive Staff",
      nameEx: "Ex Staff",
    },
    columns: {
      dcID: 5,
      name: 1,
      roles: 22,
      orders: 6,
      email: 3,
      inactiveCols: 8,
      exCols: 8,
    },
    ranges: {
      numRows: "Backend_StaffList_LastRow",
      numRowsInactive: "Backend_InactiveStaff_LastRow",
      numRowsEx: "Backend_ExStaff_LastRow",
    },
  },
  staffStats: {
    sheet: {
      id: "mainSheet",
      name: "StaffStats",
    },
    ranges: {
      basics: "StaffStats_NameMoneyOrders",
      namesVouchers: "StaffStats_Names_Vouchers",
      namesBonusExp: "StaffStats_Names_BonusExp",
      namesCargo: "StaffStats_Names_Cargo",
      typesVouchers: "StaffStats_Types_Vouchers",
      typesBonusExp: "StaffStats_Types_BonusExp",
      typesCargo: "StaffStats_Types_Cargo",
    },
    columns: {
      startVouchers: 5, // E
      numVouchers: 7, // 11 - K
      startBonusExp: 5, // E
      numBonusExp: 10, // 14 - N
      startCargo: 5, // E
      numCargo: 21, // 21 - U
    },
  },
  staffSheet: {
    id: "mainSheet",
    bonusExp: {
      name: "XP Tokens",
      colName1: 13,
      colName2: 14,
      colAmount1: 15,
      colAmount2: 16,
      colOrderID: 2,
    },
    cargo: {
      name: "Cargo",
      colName1: 14,
      colName2: 15,
      colAmount1: 16,
      colAmount2: 17,
      colOrderID: 2,
    },
  },
};

var fu = JSON.parse(JSON.stringify(fu));

var secretKey = "";

function doPost(e) {
  data = JSON.parse(e.postData.contents);
  if (data.key == null || data.key != secretKey) {
    return jsonResponse("error", "Invalid secret key");
  }
  switch (data.action) {
    case "get_worker_stats":
      return getWorkerStats();
    case "update_discord_roles":
      return updateDiscordRoles();
    case "get_discord_roles":
      return getDiscordRoles();
    case "add_order":
      return addOrder();
    case "get_worker_name":
      return getWorkerName();
    case "update_amount":
      return updateAmount();
    case "add_perms":
      return addPerms();
    case "remove_perms":
      return removePerms();
    case "rename":
      return rename();
    case "move_to_active":
      return moveToActive();
    case "move_to_inactive":
      return moveToInactive();
    case "move_to_ex":
      return moveToEx();
    case "add_to_staff_list":
      return addToStaffList();
    case "test":
      return test();
    default:
      jsonResponse("error", "unknown action");
  }
}

function getWorkerStats() {
  var filteredBase = ["Base"];
  var filteredVouchers = ["Vouchers"];
  var filteredCargo = ["Cargo"];
  var filteredBonusExp = ["BonusExp"];
  var filtered = [];

  //get workers name
  sheet = SpreadsheetApp.openById(fu["staffList"]["sheet"]["id"]).getSheetByName(fu["staffList"]["sheet"]["name"]);
  var numRows = sheet.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
  var foundRow = sheet
    .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
    .createTextFinder(data.id)
    .matchEntireCell(true)
    .findNext()
    .getRowIndex();
  var name = sheet.getRange(foundRow, fu["staffList"]["columns"]["name"], 1, 1).getValue();

  //get money and orders finished
  sheet = SpreadsheetApp.openById(fu["staffStats"]["sheet"]["id"]).getSheetByName(fu["staffStats"]["sheet"]["name"]);
  var row = sheet
    .getRange(fu["staffStats"]["ranges"]["basics"])
    .createTextFinder(name)
    .matchEntireCell(true)
    .findNext()
    .getRowIndex();
  var vals = sheet.getRange(row, 1, 1, 3).getValues();
  var basics = [vals[0][1], vals[0][2]];
  filteredBase.push(basics);
  filtered.push(filteredBase);

  //get bonus exp details
  var findNameBonusExp = sheet
    .getRange(fu["staffStats"]["ranges"]["namesBonusExp"])
    .createTextFinder(name)
    .matchEntireCell(true)
    .findNext();
  if (findNameBonusExp != null) {
    var row = findNameBonusExp.getRowIndex();

    var typesRange = sheet.getRange(fu["staffStats"]["ranges"]["typesBonusExp"]);
    var types = typesRange.getValues();

    var startCol = typesRange.getColumn();
    var numCols = types[0].length;

    var vals = sheet.getRange(row, startCol, 1, numCols).getValues();

    for (var i = 0; i < types[0].length; i++) {
      if (vals[0][i] != "" && types[0][i] != "") {
        var bonusExp = [types[0][i], vals[0][i]];
        filteredBonusExp.push(bonusExp);
      }
    }
  }
  filtered.push(filteredBonusExp);

  //get voucher details
  var findNameVouchers = sheet
    .getRange(fu["staffStats"]["ranges"]["namesVouchers"])
    .createTextFinder(name)
    .matchEntireCell(true)
    .findNext();
  if (findNameVouchers != null) {
    var row = findNameVouchers.getRowIndex();

    var typesRange = sheet.getRange(fu["staffStats"]["ranges"]["typesVouchers"]);
    var types = typesRange.getValues();

    var startCol = typesRange.getColumn();
    var numCols = types[0].length;

    var vals = sheet.getRange(row, startCol, 1, numCols).getValues();

    for (var i = 0; i < types[0].length; i++) {
      if (vals[0][i] != "" && types[0][i] != "") {
        var vouchers = [types[0][i], vals[0][i]];
        filteredVouchers.push(vouchers);
      }
    }
  }
  filtered.push(filteredVouchers);

  //get cargo details
  var findNameCargo = sheet
    .getRange(fu["staffStats"]["ranges"]["namesCargo"])
    .createTextFinder(name)
    .matchEntireCell(true)
    .findNext();
  if (findNameCargo != null) {
    var row = findNameCargo.getRowIndex();

    var typesRange = sheet.getRange(fu["staffStats"]["ranges"]["typesCargo"]);
    var types = typesRange.getValues();

    var startCol = typesRange.getColumn();
    var numCols = types[0].length;

    var vals = sheet.getRange(row, startCol, 1, numCols).getValues();

    for (var i = 0; i < types[0].length; i++) {
      if (vals[0][i] != "" && types[0][i] != "") {
        var cargo = [types[0][i], vals[0][i]];
        filteredCargo.push(cargo);
      }
    }
  }
  filtered.push(filteredCargo);

  //Logger.log(filtered);
  return jsonResponse("success", filtered);
}

function updateDiscordRoles() {
  var id = data.id;
  var rolesString = data.roles;
  try {
    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu["staffList"]["sheet"]["name"]);
    var numRows = sheet.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    var foundRow = sheet
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext()
      .getRowIndex();
    sheet.getRange(foundRow, fu["staffList"]["columns"]["roles"], 1, 1).setValue(rolesString);
    return jsonResponse("success", "roles updated");
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function getDiscordRoles() {
  var id = data.id;
  try {
    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu["staffList"]["sheet"]["name"]);
    var numRows = sheet.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    var foundRow = sheet
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext()
      .getRowIndex();
    var roles = sheet.getRange(foundRow, fu["staffList"]["columns"]["roles"], 1, 1).getValue();

    return jsonResponse("success", roles);
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function addOrder() {
  var id = data.id;
  var toAdd = data.orders;
  try {
    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu["staffList"]["sheet"]["name"]);
    var numRows = sheet.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    var foundRow = sheet
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext()
      .getRowIndex();
    var orders = sheet.getRange(foundRow, fu["staffList"]["columns"]["orders"], 1, 1).getValue().split(",");

    for (var i = 0; i < toAdd.length; i++) {
      orders.push(toAdd[i]);
    }

    var filteredOrders = orders.filter(function (order) {
      return order != "";
    });

    filteredOrders.sort();
    var ordersString = filteredOrders.join(",");

    sheet.getRange(foundRow, fu["staffList"]["columns"]["orders"], 1, 1).setValue(ordersString);

    return jsonResponse("success", "order(s) added");
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function getWorkerName() {
  var id = data.id;
  try {
    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu["staffList"]["sheet"]["name"]);
    var numRows = sheet.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    var foundRow = sheet
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext()
      .getRowIndex();
    var name = sheet.getRange(foundRow, fu["staffList"]["columns"]["name"], 1, 1).getValue();

    return jsonResponse("success", name);
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function updateAmount() {
  var order = data.order;
  var amount = data.amount;
  var name = data.name;

  var type = "";
  sheet = SpreadsheetApp.openById(fu["staffSheet"]["id"]);
  var numRows = sheet.getSheetByName("BackgroundShit").getRange("A2:B2").getValues();
  if (order.split("-")[0] == "TXE") {
    sheet = sheet.getSheetByName(fu["staffSheet"]["bonusExp"]["name"]);
    type = "bonusExp";
    numRows = numRows[0][0];
  } else if (order.split("-")[0] == "TXC") {
    sheet = sheet.getSheetByName(fu["staffSheet"]["cargo"]["name"]);
    type = "cargo";
    numRows = numRows[0][1];
  }
  Logger.log(type);
  try {
    var foundRow = sheet
      .getRange(2, fu["staffSheet"][type]["colOrderID"], numRows, 1)
      .createTextFinder(order)
      .matchEntireCell(true)
      .findNext();
    if (foundRow != null) {
      foundRow = foundRow.getRowIndex();
      var foundName = sheet.getRange(foundRow, fu["staffSheet"][type]["colName1"], 1, 1).getValue();
      if (foundName == name) {
        sheet.getRange(foundRow, fu["staffSheet"][type]["colAmount1"], 1, 1).setValue(amount);
        return jsonResponse("success", "amount updated");
      } else {
        var foundName = sheet.getRange(foundRow, fu["staffSheet"][type]["colName2"], 1, 1).getValue();
        if (foundName == name) {
          sheet.getRange(foundRow, fu["staffSheet"][type]["colAmount2"], 1, 1).setValue(amount);
          return jsonResponse("success", "amount updated");
        } else {
          return jsonResponse("error5", "you are not a worker on that order");
        }
      }
    } else {
      return jsonResponse("error5", "order not found");
    }
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function addPerms() {
  try {
    var permsToAdd = data.perms;
    var id = data.id;

    sheet = SpreadsheetApp.openById(fu.staffSheet.id);
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var tokensProtection = null;
    var cargoProtection = null;

    var staffList = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu["staffList"]["sheet"]["name"]);
    var numRows = staffList.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    var foundRow = staffList
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext()
      .getRowIndex();
    var editor = staffList
      .getRange(foundRow, fu["staffList"]["columns"]["email"], 1, 1)
      .getValue()
      .split("|")[0]
      .trim();

    for (var p in protections) {
      if (protections[p].getDescription() == "staffVouchers") {
        tokensProtection = protections[p];
      }
      if (protections[p].getDescription() == "staffCargo") {
        cargoProtection = protections[p];
      }
    }

    sheet.addEditor(editor);

    switch (permsToAdd) {
      case "both":
        tokensProtection.addEditor(editor);
        cargoProtection.addEditor(editor);
        break;
      case "tokens":
        tokensProtection.addEditor(editor);
        break;
      case "cargo":
        cargoProtection.addEditor(editor);
        break;
      default:
        break;
    }

    return jsonResponse("success", "perms added");
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function removePerms() {
  try {
    var id = data.id;
    var numRows = null;
    var foundRow = null;
    var ss = SpreadsheetApp.openById(fu.staffList.sheet.id);

    sheet = ss.getSheetByName(fu["staffList"]["sheet"]["name"]);
    numRows = ss.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
    foundRow = sheet
      .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();

    if (foundRow == null) {
      sheet = ss.getSheetByName("Inactive Staff");
      numRows = ss.getRange(fu["staffList"]["ranges"]["numRowsInactive"]).getValue();
      foundRow = sheet
        .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
        .createTextFinder(id)
        .matchEntireCell(true)
        .findNext();

      if (foundRow == null) {
        sheet = ss.getSheetByName("Ex Staff");
        numRows = ss.getRange(fu["staffList"]["ranges"]["numRowsEx"]).getValue();
        foundRow = sheet
          .getRange(2, fu["staffList"]["columns"]["dcID"], numRows, 1)
          .createTextFinder(id)
          .matchEntireCell(true)
          .findNext();
      }
    }

    var editor = sheet
      .getRange(foundRow.getRowIndex(), fu["staffList"]["columns"]["email"], 1, 1)
      .getValue()
      .split("|")[0]
      .trim();

    sheet = SpreadsheetApp.openById(fu.staffSheet.id);
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    var tokensProtection = null;
    var cargoProtection = null;

    for (var p in protections) {
      tokensProtection = protections[p].removeEditor(editor);
    }

    sheet.removeEditor(editor);

    return jsonResponse("success", "perms removed");
  } catch (err) {
    Logger.log(err);
    return jsonResponse("error", err);
  }
}

function rename() {
  var searchForVal = data.searchForVal;
  var searchForType = data.searchForType;
  var newName = data.newName;

  sheet = SpreadsheetApp.openById(fu.staffList.sheet.id).getSheetByName(fu.staffList.sheet.name);

  var numRows = sheet.getRange(fu.staffList.ranges.numRows).getValue();
  var findRow = sheet
    .getRange(2, fu.staffList.columns[searchForType], numRows, 1)
    .createTextFinder(searchForVal)
    .matchEntireCell(true)
    .findNext();
  if (findRow != null) {
    try {
      var oldName = sheet.getRange(findRow.getRowIndex(), fu.staffList.columns.name, 1, 1).getValue();

      sheet = SpreadsheetApp.openById(fu.staffSheet.id);
      var ss = null;
      var found = null;
      var currRows = null;

      sheet.getRange(findRow.getRowIndex(), fu.staffList.columns.name, 1, 1).setValue(newName);

      currRows = sheet.getSheetByName("BackgroundShit").getRange("A2:B2").getValues();

      //check XP Tokens M, N, T
      ss = sheet.getSheetByName("XP Tokens");
      found = ss
        .getRange("M2:T" + (currRows[0][0] + 1))
        .createTextFinder(oldName)
        .matchEntireCell(true)
        .matchCase(true)
        .findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      //check Cargo N, O, U
      ss = sheet.getSheetByName("Cargo");
      found = ss
        .getRange("N2:U" + (currRows[0][1] + 1))
        .createTextFinder(oldName)
        .matchEntireCell(true)
        .matchCase(true)
        .findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      currRows = sheet.getSheetByName("BackgroundShit").getRange("A6:B6").getValues();

      //check Complete XP Tokens K, L, Q
      ss = sheet.getSheetByName("Complete XP Tokens");
      found = ss
        .getRange("K2:Q" + (currRows[0][0] + 2))
        .createTextFinder(oldName)
        .matchEntireCell(true)
        .matchCase(true)
        .findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      //check Complete Cargo L, M, R
      ss = sheet.getSheetByName("Complete Cargo");
      found = ss
        .getRange("L2:R" + (currRows[0][1] + 2))
        .createTextFinder(oldName)
        .matchEntireCell(true)
        .matchCase(true)
        .findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      //check Complete Vouchers K, L, Q
      ss = sheet.getSheetByName("Complete Vouchers");
      found = ss.getRange("K2:Q3346").createTextFinder(oldName).matchEntireCell(true).matchCase(true).findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      //check Events
      ss = sheet.getSheetByName("Events");
      found = ss.getRange("C4:C200").createTextFinder(oldName).matchEntireCell(true).matchCase(true).findAll();
      for (var r in found) {
        found[r].setValue(newName);
      }

      return jsonResponse("success", oldName + " renamed to " + newName);
    } catch (err) {
      return jsonResponse("error", err);
    }
  } else {
    return jsonResponse("error", "member not found");
  }
}

function moveToActive() {
  try {
    var id = data.id;

    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id);
    var activeSheet = sheet.getSheetByName(fu.staffList.sheet.name);
    var inactiveSheet = sheet.getSheetByName(fu.staffList.sheet.nameInactive);

    var activeLastRow = sheet.getRange(fu.staffList.ranges.numRows).getValue() + 1;
    var inactiveLastRow = sheet.getRange(fu.staffList.ranges.numRowsInactive).getValue();

    var foundId = inactiveSheet
      .getRange(`E2:E${inactiveLastRow}`)
      .createTextFinder(id)
      .matchEntireCell(true)
      .findNext();
    if (foundId != null) {
      var row = foundId.getRowIndex();
      var valRange = inactiveSheet.getRange(row, 1, 1, fu.staffList.columns.inactiveCols).getValues();
      var rangeList = activeSheet
        .getRangeList([`A${activeLastRow}:F${activeLastRow}`, `J${activeLastRow}`, `V${activeLastRow}`])
        .getRanges();

      rangeList[0].setValues([
        [valRange[0][0], valRange[0][1], valRange[0][2], valRange[0][3], valRange[0][4], valRange[0][5]],
      ]);
      rangeList[1].setValue(valRange[0][6]);
      rangeList[2].setValue(valRange[0][7]);

      inactiveSheet.deleteRow(foundId.getRowIndex());

      return jsonResponse("success", "worker moved to active");
    } else {
      return jsonResponse("error", "worker not found");
    }
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function moveToInactive() {
  try {
    var id = data.id;

    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id);
    var activeSheet = sheet.getSheetByName(fu.staffList.sheet.name);
    var inactiveSheet = sheet.getSheetByName(fu.staffList.sheet.nameInactive);

    var activeLastRow = sheet.getRange(fu.staffList.ranges.numRows).getValue();
    var inactiveLastRow = sheet.getRange(fu.staffList.ranges.numRowsInactive).getValue() + 1;

    var foundId = activeSheet.getRange(`E2:E${activeLastRow}`).createTextFinder(id).matchEntireCell(true).findNext();
    if (foundId != null) {
      var row = foundId.getRowIndex();
      var rangeList = activeSheet.getRangeList([`A${row}:F${row}`, `J${row}`, `V${row}`]).getRanges();
      var valRange = rangeList[0].getValues();
      valRange[0].push(rangeList[1].getValue(), rangeList[2].getValue());

      inactiveSheet.getRange(inactiveLastRow, 1, 1, fu.staffList.columns.inactiveCols).setValues(valRange);

      activeSheet.deleteRow(row);

      return jsonResponse("success", "worker moved to inactive");
    } else {
      return jsonResponse("error", "worker not found");
    }
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function moveToEx() {
  try {
    var id = data.id;
    var reason = data.reason;
    var ss = SpreadsheetApp.openById(fu.staffList.sheet.id);

    sheet = ss.getSheetByName(fu.staffList.sheet.name);
    var lastRow = ss.getRange(fu.staffList.ranges.numRows).getValue();

    var foundId = sheet.getRange(`E2:E${lastRow}`).createTextFinder(id).matchEntireCell(true).findNext();
    if (foundId == null) {
      sheet = ss.getSheetByName(fu.staffList.sheet.nameInactive);
      lastRow = ss.getRange(fu.staffList.ranges.numRowsInactive).getValue();
      foundId = sheet.getRange(`E2:E${lastRow}`).createTextFinder(id).matchEntireCell(true).findNext();
    }

    var exSheet = ss.getSheetByName(fu.staffList.sheet.nameEx);
    var exLastRow = ss.getRange(fu.staffList.ranges.numRowsEx).getValue() + 1;
    var row = foundId.getRowIndex();
    var range;

    if (sheet.getSheetName() == "Staff List") {
      var rangeList = sheet.getRangeList([`A${row}:F${row}`, `J${row}`]).getRanges();
      range = rangeList[0].getValues();
      range[0].push(rangeList[1].getValue(), reason);
    } else if (sheet.getSheetName() == "Inactive Staff") {
      range = sheet.getRange(`A${row}:G${row}`).getValues();
      range[0].push(reason);
    }

    exSheet.getRange(exLastRow, 1, 1, fu.staffList.columns.exCols).setValues(range);
    sheet.deleteRow(row);

    return jsonResponse("success", "worker moved to ex");
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function addToStaffList() {
  try {
    sheet = SpreadsheetApp.openById(fu.staffList.sheet.id);
    var row = sheet.getSheetByName("Backend").getRange(fu.staffList.ranges.numRows).getValue() + 1;
    var rangeList = sheet
      .getSheetByName(fu.staffList.sheet.name)
      .getRangeList([`A${row}:E${row}`, `J${row}`])
      .getRanges();
    rangeList[0].setValues([[data.name, data.igID, `${data.email} | `, data.rank, data.dcID]]);
    rangeList[1].setValue(data.date);

    return jsonResponse("success", "added to the list");
  } catch (err) {
    return jsonResponse("error", err);
  }
}

function test() {
  var numRows = null;
  var foundRow = null;
  sheet = SpreadsheetApp.openById(fu.staffList.sheet.id);

  var staffList = sheet.getSheetByName(fu["staffList"]["sheet"]["name"]);
  numRows = staffList.getRange(fu["staffList"]["ranges"]["numRows"]).getValue();
  Logger.log(numRows);
  numRows = sheet.getRange(fu.staffList.ranges.numRows).getValue();
  Logger.log(numRows);
}

function jsonResponse(state, message) {
  return ContentService.createTextOutput(JSON.stringify({ state: state, message: message }));
}
