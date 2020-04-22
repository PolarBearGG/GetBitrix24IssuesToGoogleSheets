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
        .addItem('Get Data', 'refresh')
        .addToUi();
}
var token = "";
var userId = "";
var cnt = 5000;
function getTasks() {
    //------------------------------------------------------------------------------------------------------------------------------//
    var start;
    var xmlARR = [];
    for (start = 0; start <= cnt; start = start + 50) {
        var FeedURL = token + "tasks.task.list.xml?start=" + start + "&order[CREATED_DATE]=desc";
        // Generate 2d/md array / rows export based on requested columns and feed
        var exportRows = []; // hold all the rows that are generated to be pasted into the sheet
        var XMLFeedURL = FeedURL;
        var feedContent = UrlFetchApp.fetch(XMLFeedURL).getContentText(); // get the full feed content
        var feedItems = XmlService.parse(feedContent).getRootElement().getChild('result').getChild('tasks').getChildren('item'); // get all items in the feed
        var next = XmlService.parse(feedContent).getRootElement().getChildText('next');
      Logger.log(feedItems.length)
      
       var nodeArray = ["id", "title", "createdDate", "closedDate", "timeEstimate", "responsible", "timeSpentInLogs"];
       if(next){
        for (var x = 0; x < feedItems.length; x++) {
            var currentFeedItem = feedItems[x];
            var singleItemArray = [];
            for (var y = 0; y < nodeArray.length; y++) {
               
             
              
              if(nodeArray[y]==="responsible") {
               if (currentFeedItem.getChild(nodeArray[y]).getChildren("name")) {
                    singleItemArray.push(currentFeedItem.getChild(nodeArray[y]).getChildText("name"));
                } else {
                    singleItemArray.push("null");
                }
              } else {
              
              if (currentFeedItem.getChild(nodeArray[y])) {
                    singleItemArray.push(currentFeedItem.getChildText(nodeArray[y]));
                } else {
                    singleItemArray.push("null");
                }
              }
            }
            exportRows.push(singleItemArray);
        }
        xmlARR.push(exportRows);
       }else { break;}} 
    var GoogleSheetsFile = SpreadsheetApp.getActiveSpreadsheet();
    var GoogleSheetsPastePage = GoogleSheetsFile.getSheetByName('bitrixQ1');
    Logger.log([nodeArray]);
    GoogleSheetsPastePage.getRange(1, 1, 1, nodeArray.length).setValues([nodeArray]);
    GoogleSheetsPastePage.getDataRange().offset(1, 0).clearContent();
    for (var i = 0; i < xmlARR.length; i++) {
        GoogleSheetsPastePage.getRange(GoogleSheetsPastePage.getLastRow() + 1, 1, xmlARR[i].length, xmlARR[i][1].length).setValues(xmlARR[i]);
    }
}
function refresh() {
   
    getTasks();
    Logger.log('Done! ');
}
