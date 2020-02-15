/**
 * ###########################################################################
 * # Name: Bitrix24 Automation                                               #
 * # Description: This script let's you connect to Bitrix24 CRM and retrieve #
 * #              its data to populate a Google Spreadsheet.                 #
 * # Date: February 15th, 2020                                               #
 * # Author: Korolyk Vitaliy                                                 #
 * # Modified by:                                                            #
 * # Detail of the turorial:                                                 #
 * ###########################################################################
 */

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Bitrix24 Connector')
        .addItem('Set Bitrix24 token', 'setBitrix24Token')
        .addItem('Set User ID', 'setBitrix24UserId')
        .addItem('Set Issues Count', 'setIssuesCount')
        .addSeparator()
        .addItem('Get Data', 'refresh')
        .addSeparator()
        .addItem('Reset Credentials', 'reset')
        .addToUi();
}
var token = documentProperties.getProperty('BITRIX24_TOKEN');
var userId = documentProperties.getProperty('BITRIX24_USER_ID');
var cnt = documentProperties.getProperty('ISSUES_COUNT');


function setIssuesCount() {
    var ui = SpreadsheetApp.getUi();


    var result = ui.prompt(
        'Issues Count',
        'Please enter issues count:',
        ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
        // User clicked "OK".
        documentProperties.setProperty("ISSUES_COUNT", text);
    } else if (button == ui.Button.CANCEL) {
        // User clicked "Cancel".
    } else if (button == ui.Button.CLOSE) {
        // User clicked X in the title bar.

    }

}





function setBitrix24Token() {
    var ui = SpreadsheetApp.getUi();


    var result = ui.prompt(
        'Bitrix24 Token',
        'Please enter your Bitrix24 Token:',
        ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
        // User clicked "OK".
        documentProperties.setProperty("BITRIX24_TOKEN", text);
    } else if (button == ui.Button.CANCEL) {
        // User clicked "Cancel".
    } else if (button == ui.Button.CLOSE) {
        // User clicked X in the title bar.

    }

}

function setBitrix24UserId() {
    var ui = SpreadsheetApp.getUi(); // Same variations.

    var result = ui.prompt(
        'Bitrix24 User ID',
        'Please enter your Bitrix24 User ID:',
        ui.ButtonSet.OK_CANCEL);

    // Process the user's response.
    var button = result.getSelectedButton();
    var text = result.getResponseText();
    if (button == ui.Button.OK) {
        // User clicked "OK".
        documentProperties.setProperty("BITRIX24_USER_ID", text);
    } else if (button == ui.Button.CANCEL) {
        // User clicked "Cancel".
    } else if (button == ui.Button.CLOSE) {
        // User clicked X in the title bar.

    }

}


function reset() {

    documentProperties.deleteProperty("BITRIX24_TOKEN");
    documentProperties.deleteProperty("BITRIX24_USER_ID");
    documentProperties.deleteProperty("ISSUES_COUNT");

}

function getTasks() {


    //------------------------------------------------------------------------------------------------------------------------------//
    var start;
    var xmlARR = [];
    for (start = 0; start <= cnt; start = start + 50) {
        var FeedURL = token + "task.item.list.xml?start=" + start + "&O[CREATED_DATE]=&F[RESPONSIBLE_ID]=" + userId + "&P[]=";
        // Generate 2d/md array / rows export based on requested columns and feed
        var exportRows = []; // hold all the rows that are generated to be pasted into the sheet
        var XMLFeedURL = FeedURL;
        var feedContent = UrlFetchApp.fetch(XMLFeedURL).getContentText(); // get the full feed content
        var feedItems = XmlService.parse(feedContent).getRootElement().getChild('result').getChildren('item'); // get all items in the feed
        var nodeArray = ["ID", "TITLE", "CREATED_DATE", "CLOSED_DATE", "TIME_ESTIMATE", "PRIORITY", "GROUP_ID", "DESCRIPTION", "RESPONSIBLE_ID", "COMMENTS_COUNT"];
        for (var x = 0; x < feedItems.length; x++) {
            var currentFeedItem = feedItems[x];
            var singleItemArray = [];
            for (var y = 0; y < nodeArray.length; y++) {
                if (currentFeedItem.getChild(nodeArray[y])) {
                    singleItemArray.push(currentFeedItem.getChildText(nodeArray[y]));
                } else {
                    singleItemArray.push("null");
                }
            }
            exportRows.push(singleItemArray);
        }
        xmlARR.push(exportRows);
    }




    var GoogleSheetsFile = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1LWywk2SIrJ22lGXwIcycN2dBehkCSV-y0zBhoVSGsb0/edit#gid=0");
    var GoogleSheetsPastePage = GoogleSheetsFile.getSheetByName('Bitrix');
    Logger.log([nodeArray]);
    GoogleSheetsPastePage.getRange(1, 1, 1, nodeArray.length).setValues([nodeArray]);

    GoogleSheetsPastePage.getDataRange().offset(1, 0).clearContent();
    for (var i = 0; i < xmlARR.length; i++) {

        GoogleSheetsPastePage.getRange(GoogleSheetsPastePage.getLastRow() + 1, 1, xmlARR[i].length, xmlARR[i][1].length).setValues(xmlARR[i]);
    }
}









function refresh() {
    if (!documentProperties.getProperty("BITRIX24_TOKEN")) {
        setBitrix24Token();
    }
    if (!documentProperties.getProperty("BITRIX24_USER_ID")) {
        setBitrix24UserId();
    }
    getTasks();
    Logger.log('Done! ');
}