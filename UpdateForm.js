function updateForm(e) {
  var range = e.range;
  if (range.getA1Notation() == "G17" && range.getSheet().getName() == "BackgroundShit" && range.getValue() == "Yes") {
    var sheet = SpreadsheetApp.openById("mainSheet").getSheetByName("BackgroundShit");

    var form = FormApp.openById("formID");
    var items = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);

    /* UPDATE VOUCHERS */
    var item = items[1].asMultipleChoiceItem();
    var bonusExp = [];
    var bonusExpRange = sheet.getRange("VoucherInfo").getDisplayValues();
    for (var i = 0; i < bonusExpRange.length; i++) {
      var temp = bonusExpRange[i][0] + " | " + bonusExpRange[i][1] + " | Limit " + bonusExpRange[i][2];
      bonusExp.push(item.createChoice(temp));
    }
    item.setChoices(bonusExp);

    /* UPDATE CARGO */
    var item = items[2].asMultipleChoiceItem();
    var cargo = [];
    var cargoRange = sheet.getRange("CargoInfo").getDisplayValues();
    for (var i = 0; i < cargoRange.length; i++) {
      var temp = cargoRange[i][0] + " | " + cargoRange[i][1] + " | Limit " + cargoRange[i][2];
      cargo.push(item.createChoice(temp));
    }
    item.setChoices(cargo);

    range.setValue("No");
  }
}
