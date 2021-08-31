/* Global variables */
var ss = SpreadsheetApp.openById("mainSheet");
var staffList = SpreadsheetApp.openById("staffList");
var sBackend = staffList.getSheetByName("Backend");
var sResponses = ss.getSheetByName("Responses");
var sVouchers = ss.getSheetByName("XP Tokens");
var sCargo = ss.getSheetByName("Cargo");
var sShit = ss.getSheetByName("BackgroundShit");
var sStats = ss.getSheetByName("Stats");
var sBroken = ss.getSheetByName("Broken Responses");

function responseFormSubmit(e) {
  var workbitch = sResponses.getLastRow() - 1;

  if (workbitch > 0) {
    for (bitch = 0; bitch < workbitch; bitch++) {
      try {
        /* Load order details */
        var order = sResponses.getRange(2, 1, 1, 13).getValues();
        order = order[0];
        var orderType = order[4];
        order.splice(4, 1);

        /* Set priority to the right format */
        if (order[9] == "High Priority") {
          order[9] = "High";
        } else {
          order[9] = "Normal";
        }

        /* Delete empty indexes based on the type */
        if (orderType == "Cargo") {
          order.splice(4, 2);
        } else {
          order.splice(6, 3);
        }

        /* Get price and limit info */
        var priceDetails = order[4].split("|");
        for (var p in priceDetails) {
          priceDetails[p] = priceDetails[p].trim();
        }
        priceDetails[1] = parseInt(priceDetails[1].replace("$", "").replace(",", ""));
        priceDetails[2] = parseInt(priceDetails[2].split(" ")[1].trim().replace(",", ""));
        order[4] = priceDetails[0]; // Set the Voucher/Cargo type to the right format

        /* Set storage to the right format */
        if (orderType == "Cargo") {
          var index = order[6].indexOf("(");
          if (index > 0) {
            order[6] = order[6].substring(0, index - 1);
          } else if (order[6] == "Any Public Storage [+10%]") {
            order[6] = "Any Public Storage";
          } else {
            order[6] = "";
          }
        }

        /* Check if order would go over the limit */
        if (orderType == "Bonus EXP Tokens") {
          var orderOK = limitCheck(
            orderType,
            priceDetails,
            order[2],
            "Stats_OrderCheck_VType",
            "Stats_OrderCheck_VIDs",
            order[5]
          );
        } else {
          var orderOK = limitCheck(
            orderType,
            priceDetails,
            order[2],
            "Stats_OrderCheck_CType",
            "Stats_OrderCheck_CIDs",
            order[5]
          );
        }
        if (orderOK[0] == "error") {
          if (orderOK[1] == "over limit") {
            sResponses.deleteRow(2);
            sendWebhook(order[5] + " " + order[4] + " order from: " + order[3] + " deleted", "reason: over the limit");
            throw new Error("order over limit");
          }
          if (orderOK[1] == "over type limit") {
            sResponses.deleteRow(2);
            sendWebhook(
              order[5] + " " + order[4] + " order from: " + order[3] + " deleted",
              "reason: over the type limit"
            );
            throw new Error("order over type limit");
          }
        }

        /* Discounts / codes / price formula */
        var codeDiscount = 0;
        var vPriceFormula =
          '=(IF(R[0]C[-5]<>"",IF(R[0]C[-4]<>"",R[0]C[-4]*' +
          priceDetails[1] +
          '*(100%-R[0]C[-1])*IF(R[0]C[-3]="High",1.35,1),""),""))';
        var cPriceFormula =
          '=IF(R[0]C[-6]<>"",IF(R[0]C[-5]<>"",R[0]C[-5]*' +
          priceDetails[1] +
          '*(100%-R[0]C[-1])*IF(R[0]C[-3]="High",1.35,1)*IF(R[0]C[-4]<>"",100%+VLOOKUP(R[0]C[-4],StoragePercent,2,FALSE),1),""),"")';

        if (order[order.length - 1] != "") {
          var codeRange = sBackend.getRange("Backend_PromoCodes");
          var codeFind = codeRange
            .createTextFinder(order[order.length - 1])
            .matchEntireCell(true)
            .findNext();
          if (codeFind != null) {
            var codeRow = codeFind.getRow();
            var codeCol = codeFind.getColumn();
            var codeInfo = sBackend.getRange(codeRow, codeCol, 1, 4).getValues()[0];

            if (codeInfo[3] > 0) {
              sBackend.getRange(codeRow, codeCol + 3, 1, 1).setValue(codeInfo[3] - 1);

              if (codeInfo[1] == "Discount") {
                codeDiscount = codeInfo[2] * 100;
                order[order.length - 2] += " | Discount Code |";
              } else if (codeInfo[1] == "Free Priority") {
                order[order.length - 3] = "High";
                vPriceFormula =
                  '=(IF(R[0]C[-5]<>"",IF(R[0]C[-4]<>"",R[0]C[-4]*' + priceDetails[1] + '*(100%-R[0]C[-1]),""),""))';
                cPriceFormula =
                  '=IF(R[0]C[-6]<>"",IF(R[0]C[-5]<>"",R[0]C[-5]*' +
                  priceDetails[1] +
                  '*(100%-R[0]C[-1])*IF(R[0]C[-4]<>"",100%+VLOOKUP(R[0]C[-4],StoragePercent,2,FALSE),1),""),"")';
                order[order.length - 2] += " | Free Priority |";
              } else if (codeInfo[1] == "Free Storage") {
                if (orderType == "Cargo") {
                  cPriceFormula =
                    '=IF(R[0]C[-6]<>"",IF(R[0]C[-5]<>"",R[0]C[-5]*' +
                    priceDetails +
                    '*(100%-R[0]C[-1])*IF(R[0]C[-3]="High",1.35,1),""),"")';
                  order[order.length - 2] += " | Free Storage |";
                }
              }
            }
          }
        }
        var discount = 0;
        var discounts = sShit.getRange("F21:G22").getValues();
        if (orderType == "Bonus EXP Tokens" && discounts[0][1] == "Yes") {
          discount = discounts[0][0] * 100;
        } else if (orderType == "Cargo" && discounts[1][1] == "Yes") {
          discount = discounts[1][0] * 100;
        }

        var loyalDiscount = 0;
        var loyalID = sShit.getRange("LoyalCustomers").createTextFinder(order[2]).matchEntireCell(true).findNext();
        if (loyalID != null) {
          loyalDiscount = 5;
        }

        var totalDiscount = discount + codeDiscount + loyalDiscount;
        if (totalDiscount > 0) {
          order[order.length - 1] = totalDiscount + "%";
        } else {
          order[order.length - 1] = "";
        }

        /* push price formula */
        if (orderType == "Bonus EXP Tokens") {
          order.push(vPriceFormula);
        } else {
          order.push(cPriceFormula);
        }

        /* add order ID, push status */
        var nextIDs = sShit.getRange("A4:B4").getValues()[0];
        if (orderType == "Bonus EXP Tokens") {
          order.splice(1, 0, "TXE-" + nextIDs[0]);
          sShit.getRange("A4").setValue(nextIDs[0] + 1);
        } else {
          order.splice(1, 0, "TXC-" + nextIDs[1]);
          sShit.getRange("B4").setValue(nextIDs[1] + 1);
        }
        order.push("Queued", "", "", "", "");
        /*if(orderType=="Vouchers"){
            order.push("=IF(ISERROR(SUM(R[0]C[-2],R[0]C[-1])/R[0]C[-10]),\"\",SUM(R[0]C[-2],R[0]C[-1])/R[0]C[-10])","");
          }else{
            order.push("=IF(ISERROR(SUM(R[0]C[-2],R[0]C[-1])/R[0]C[-11]),\"\",SUM(R[0]C[-2],R[0]C[-1])/R[0]C[-11])","");
          }*/
        order.push("", "", "", "", "Yes");

        /*Logger.log(order);
          for each(var v in order){
            Logger.log(v);
          }
          throw new Error("testing mid operation haha");*/

        /* add order to the right sheet */
        var arr = [];
        arr[0] = order;

        var nextOrderRow = sShit.getRange("A2:B2").getValues()[0];

        var protsRange = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);
        var protsSheet = ss.getProtections(SpreadsheetApp.ProtectionType.SHEET);
        var rangeProt = null;
        var sheetProt = null;

        if (orderType == "Bonus EXP Tokens") {
          var lastRow = sVouchers.getLastRow();
          var lastCol = sVouchers.getLastColumn();

          sVouchers.getRange(nextOrderRow[0] + 2, 1, 1, 21).setValues(arr);
          sVouchers.insertRowAfter(lastRow);
          sVouchers
            .getRange(lastRow, 1, 1, lastCol)
            .copyTo(sVouchers.getRange(lastRow + 1, 1, 1, lastCol), { contentsOnly: false });

          for (var p in protsRange) {
            var val = protsRange[p];
            if (val.getDescription() == "staffVouchers") {
              rangeProt = val;
            }
          }

          for (var p in protsSheet) {
            var val = protsSheet[p];
            if (val.getDescription() == "vouchers_all") {
              sheetProt = val;
            }
          }

          var newRange = sVouchers.getRange("L2:W" + (lastRow + 1));
          rangeProt.setRange(newRange);
          sheetProt.setUnprotectedRanges([newRange]);
        } else {
          var lastRow = sCargo.getLastRow();
          var lastCol = sCargo.getLastColumn();

          sCargo.getRange(nextOrderRow[1] + 2, 1, 1, 22).setValues(arr);
          sCargo.insertRowAfter(lastRow);
          sCargo
            .getRange(lastRow, 1, 1, lastCol)
            .copyTo(sCargo.getRange(lastRow + 1, 1, 1, lastCol), { contentsOnly: false });

          for (var p in protsRange) {
            var val = protsRange[p];
            if (val.getDescription() == "staffCargo") {
              rangeProt = val;
            }
          }

          for (var p in protsSheet) {
            var val = protsSheet[p];
            if (val.getDescription() == "cargo_all") {
              sheetProt = val;
            }
          }

          var newRange = sCargo.getRange("M2:X" + (lastRow + 1));
          rangeProt.setRange(newRange);
          sheetProt.setUnprotectedRanges([newRange]);
        }

        sResponses.deleteRow(2);
      } catch (error) {
        try {
          var order = sResponses.getRange(2, 1, 1, 13).getValues();
          var lastRow = sBroken.getLastRow() + 1;
          sBroken.getRange(lastRow, 1, 1, 13).setValues(order);
          sResponses.deleteRow(2);
          Logger.log(error);
        } catch (err) {
          Logger.log(err);
        }
      }
    }
  }
}
//------------------------------------------------------------------------------------------------------------------------------------------//

function limitCheck(_orderType, _priceDetails, _playerID, _orderCheckTypeRange, _orderCheckIDRange, _orderAmount) {
  var limitOK = false;
  var typeLimitOK = true;
  /*if(_orderType == "Cargo"){
    var typeLimitOK = true;
  }else{
    var typeLimitOK = false;
  }*/

  /* Normal limit check */
  var currOrderCol = sStats
    .getRange(_orderCheckTypeRange)
    .createTextFinder(_priceDetails[0])
    .matchEntireCell(true)
    .findNext();
  var currOrderRow = sStats.getRange(_orderCheckIDRange).createTextFinder(_playerID).matchEntireCell(true).findNext();
  var currOrdered = 0;
  var diffTypes = [];

  if (currOrderRow != null && currOrderCol != null) {
    currOrdered = sStats.getRange(currOrderRow.getRow(), currOrderCol.getColumn()).getValue();
  } else {
    currOrdered = 0;
  }
  if (currOrdered + _orderAmount <= _priceDetails[2]) {
    limitOK = true;
  } else {
    return ["error", "over limit"];
  }

  /* Type limit check */
  /*if(limitOK == true && _orderType=="Bonus EXP Tokens"){
    if(currOrderRow!=null && sStats.getRange("M"+currOrderRow.getRow()).getValue()>=2){
    var orderedStuffRange = sStats.getRange(currOrderRow.getRow(), 3, 1, 9).getValues();
    var helpMEcounter = 3;
    orderedStuffRange[0].forEach(
      i=>{
        if(i!=""){
          diffTypes.push(sStats.getRange(137, helpMEcounter).getValue());
        }
        helpMEcounter++;
      });
      if(diffTypes.indexOf(_priceDetails[0]) < 0){
        return ["error","over type limit"];
      }else{
        typeLimitOK = true;
      }
    }
  }*/

  if (limitOK == true && typeLimitOK == true) {
    return ["success", "everything is ok"];
  } else {
    return ["error", "unknown error"];
  }
}

function sendWebhook(name, value) {
  var POST_URL = "dcWebHook";
  var items = [];
  items.push({
    name: name,
    value: value,
    inline: false,
  });
  var options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
    },
    payload: JSON.stringify({
      embeds: [
        {
          title: "",
          fields: items,
          footer: {
            text: "",
          },
        },
      ],
    }),
  };

  UrlFetchApp.fetch(POST_URL, options);
}
