function moveCompleteCancelled(e) {
  var range = e.range;
  if (range.getA1Notation() == "G26" && range.getSheet().getName() == "BackgroundShit" && range.getValue() == "Yes") {
    var debug = false;
    var ss = SpreadsheetApp.openById("mainSheet");
    var sheet_xp_tokens = ss.getSheetByName("XP Tokens");
    var sheet_xp_tokens_complete = ss.getSheetByName("Complete XP Tokens");
    var sheet_cargo = ss.getSheetByName("Cargo");
    var sheet_cargo_complete = ss.getSheetByName("Complete Cargo");
    var sheet_background_shit = ss.getSheetByName("BackgroundShit");
    var things = {
      columns_xp: 25,
      columns_cargo: 26,
      current_xp_orders: "A2",
      current_cargo_orders: "B2",
      status_xp: 11,
      status_cargo: 12,
    };
    var filtered_xp = [];
    var filtered_cargo = [];
    var to_move_xp = 0;
    var to_move_cargo = 0;
    var rows_to_delete_xp = [];
    var rows_to_delete_cargo = [];
    var current_row = 2;
    var placeholder = [];

    var vOrders = sheet_background_shit.getRange(things.current_xp_orders).getValue();
    var cOrders = sheet_background_shit.getRange(things.current_cargo_orders).getValue();
    if (vOrders == 0) vOrders = 1;
    if (cOrders == 0) cOrders = 1;

    var bonus_xp_range_values = sheet_xp_tokens.getRange(2, 1, vOrders, things.columns_xp).getValues();
    var cargo_range_values = sheet_cargo.getRange(2, 1, cOrders, things.columns_cargo).getValues();

    for (row of bonus_xp_range_values) {
      if (row[things.status_xp] == "Complete" || row[things.status_xp] == "Cancelled") {
        placeholder = [
          row[0], // Timestamp
          row[1], // Order ID
          row[2], // In-game Name
          row[3], // In-game ID
          row[4], // Discord ID
          row[5], // Voucher Type
          row[6], // Amount
          row[7], // Priority
          row[9], // Discount
          row[10], // Total Price
          row[12], // Worker #1
          row[13], // Worker #2
          row[14], // Worker #1 Amount
          row[15], // Worker #2 Amount
          row[17], // Worker #1 Cut
          row[18], // Worker #2 Cut
          row[19], // Collector
          row[8], // Additional Information
          row[things.status_xp].toUpperCase(), // Status
          Utilities.formatDate(new Date(), "GMT", "YYY-MM-dd"), // Date
          row[22], // Staff Additional Info
        ];
        filtered_xp.push(placeholder);
        to_move_xp++;
        rows_to_delete_xp.push(current_row);
      }

      current_row++;
    }
    current_row = 2;

    for (row of cargo_range_values) {
      if (row[things.status_cargo] == "Complete" || row[things.status_cargo] == "Cancelled") {
        placeholder = [
          row[0], // Timestamp
          row[1], // Order ID
          row[2], // In-game Name
          row[3], // In-game ID
          row[4], // Discord ID
          row[5], // Cargo Type
          row[6], // Amount
          row[7], // Storage
          row[8], // Priority
          row[10], // Discount
          row[11], // Total Price
          row[13], // Worker #1
          row[14], // Worker #2
          row[15], // Worker #1 Amount
          row[16], // Worker #2 Amount
          row[18], // Worker #1 Cut
          row[19], // Worker #2 Cut
          row[20], // Collector
          row[9], // Additional Information
          row[things.status_cargo].toUpperCase(), // Status
          Utilities.formatDate(new Date(), "GMT", "YYY-MM-dd"), // Date
          row[23], // Staff Additional Info
        ];
        filtered_cargo.push(placeholder);
        to_move_cargo++;
        rows_to_delete_cargo.push(current_row);
      }

      current_row++;
    }

    if (to_move_xp > 0) {
      var last_row = sheet_xp_tokens_complete.getLastRow();
      sheet_xp_tokens_complete.insertRowsAfter(last_row, to_move_xp);

      var range_start = last_row + 1;
      sheet_xp_tokens_complete.getRange(range_start, 1, to_move_xp, 21).setValues(filtered_xp);

      //delete
      for (var i = rows_to_delete_xp.length - 1; i > -1; i--) {
        if (!debug) {
          //sheet_xp_tokens.deleteRow(rows_to_delete_xp[i]);
        } else {
          Logger.log(rows_to_delete_xp[i]);
        }
      }
    }

    if (to_move_cargo > 0) {
      var last_row = sheet_cargo_complete.getLastRow();
      sheet_cargo_complete.insertRowsAfter(last_row, to_move_cargo);

      var range_start = last_row + 1;
      sheet_cargo_complete.getRange(range_start, 1, to_move_cargo, 22).setValues(filtered_cargo);

      //delete
      for (var i = rows_to_delete_cargo.length - 1; i > -1; i--) {
        if (!debug) {
          //sheet_cargo.deleteRow(rows_to_delete_cargo[i]);
        } else {
          Logger.log(rows_to_delete_cargo[i]);
        }
      }
    }

    range.setValue("No");
  }
}
