var vars = {
	"XP Tokens":"X",
    "Cargo":"Y",
};

var vars = JSON.parse(JSON.stringify(vars));

function updateDateCollected(e) {
    var range = e.range;
    let sheetName = range.getSheet().getName();
    if( (sheetName == "XP Tokens" || sheetName == "Cargo") && e.value == "Awaiting Client Collection" ){
        let sheet = e.source.getSheetByName(sheetName);
        let row = range.getRow();
        const date = new Date(Date.now());
        //let currDate = `${date.getDate()}/${date.getMonth()+1}/${date.getFullYear()} ${date.getHours()}:${date.getMinutes()}`;
        sheet.getRange(`${vars[sheetName]}${row}`).setValue(date);
    }
  }