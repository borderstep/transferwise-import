/////////////////////////////////////////////////////////////////////////
// YOUR API KEY GOES BETWEEN THE QUOTATION MARKS
var API = "xxx-your-transferwise-readOnly-api-key-goes-here-xxx"

// STATEMENT BEGIN DATE | YYYY-MM-DD
var beginDate = "2019-04-01";

// STATEMENT END DATE | YYYY-MM-DD
var endDate   = "2019-04-30";

// SELECT CELL IN YOUR SHEET AND CHOOSE YOUR OPTION FROM THE SHEETS' MENU
/////////////////////////////////////////////////////////////////////////
// The below code fetches the profile id with the API key, then fetches
// the borderless account id thanks to the profile id and then fetches
// the relevant transaction data and sprays it into the sheet. Voilà.
////////////////////////////
// PREPARE HEADERS
var auth = "Bearer ".concat(API);
var headers = {
"Authorization" : auth
};
var params = {
"method" : "GET",
"headers" : headers
};
////////////////////////////
// ACCESS PROFILE
var idFetch = UrlFetchApp.fetch("https://api.transferwise.com/v1/profiles", params);
var idResponse = [idFetch.getContentText()];
var idJSONObject = JSON.parse(idResponse);
var id = idJSONObject[0].id;
////////////////////////////
// ACCESS BORDERLESS ACCOUNT
var idFetchB = UrlFetchApp.fetch("https://api.transferwise.com/v1/borderless-accounts?profileId=" + id, params);
var idResponseB = [idFetchB.getContentText()];
var idJSONObjectB = JSON.parse(idResponseB);
var idB = idJSONObjectB[0].id;
////////////////////////////
// PREPARE RAW DATA
var borderlessUrl = "https://api.transferwise.com/v1/borderless-accounts/" + idB + "/statement.json?currency=EUR&intervalStart=" + beginDate + "T00:00:00.000Z&intervalEnd=" + endDate + "T23:59:59.999Z"
var response = UrlFetchApp.fetch(borderlessUrl, params);
var statement = [response.getContentText()];
var JSONObject = JSON.parse(statement);
var sheet = SpreadsheetApp.getActiveSheet();
var tx = JSONObject.transactions.reverse();
var txLength = tx.length;
////////////////////////////
// MOTHER-FUNCTION
function DebitsCredits(){
    const firstActive = sheet.getActiveCell();
	borderlessDebits();
    firstActive.offset(0, 5).activate();
	borderlessCredits();
    firstActive.activate();
}
////////////////////////////
// SPRAY DEBITS
function borderlessDebits() {
    const firstActive = sheet.getActiveCell();
    for (var i = 0; i < txLength; i++) {
      if (tx[i].type == "DEBIT") {
        var active = sheet.getActiveCell();
// DATE
        date = (tx[i].date).toString();
        formattedDate = date.slice(5,7) + "/" + date.slice(8,10) + "/" + date.slice(0,4);
        active.setValue(formattedDate);
// AMOUNT
        active.offset(0, 1).setValue([tx[i].amount.value]*(-1));
// DESCRIPTION
        var description;
        if (tx[i].details.type == "TRANSFER") {
          description = tx[i].details.description.concat(" - ", tx[i].details.paymentReference);
        } else if ((tx[i].details.type == "CARD")) {
          description = tx[i].details.merchant.name.concat(" ", tx[i].details.merchant.category);
        } else {
          description = "";
        }
        active.offset(0, 2).setValue(description);
// NEXT LINE
        active.offset(1, 0).activate();
      }
    }
    firstActive.activate();
}
////////////////////////////
// SPRAY CREDITS
function borderlessCredits() {
    const firstActive = sheet.getActiveCell();
    for (var i = 0; i < txLength; i++) {
      if (tx[i].type == "CREDIT") {
        var active = sheet.getActiveCell();
// DATE
        date = (tx[i].date).toString();
        formattedDate = date.slice(5,7) + "/" + date.slice(8,10) + "/" + date.slice(0,4);
        active.setValue(formattedDate);
// AMOUNT
        active.offset(0, 1).setValue([tx[i].amount.value]);
//DESCRIPTION
        var description;
        if (tx[i].details.type == "MONEY_ADDED") {
          description = tx[i].details.description.concat(" - ", tx[i].referenceNumber);
        } else if ((tx[i].details.type == "CARD")) {
          description = tx[i].details.merchant.name.concat(" ", tx[i].details.merchant.category);
        } else {
          description = "";
        }
        active.offset(0, 2).setValue(description);
// NEXT LINE
        active.offset(1, 0).activate();
      }
    }
    firstActive.activate();
}
////////////////////////////
// ADD RELEVANT MENU ENTIRES
function onOpen() {
	var ui = SpreadsheetApp.getUi();
	var menu = ui.createMenu("Import TransferWise");
	var item = menu.addItem("Debits & Credits", 'DebitsCredits');
	var item2 = menu.addItem("Debits", 'borderlessDebits');
	var item3 = menu.addItem("Credits", 'borderlessCredits');
	item.addToUi();
};
onOpen();
////////////////////////////
// END
