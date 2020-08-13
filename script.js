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
var token = "";
var userId = "";
var cnt = 5000;
function getTasks() {
    //------------------------------------------------------------------------------------------------------------------------------//
    var start;
    var xmlARR = [];
    for (start = 0; start <= cnt; start = start + 50) {
        var FeedURL = token + "tasks.task.list.xml?start=" + start + "&order[CLOSED_DATE]=desc&filter[RESPONSIBLE_ID]=" + userId;
        // Generate 2d/md array / rows export based on requested columns and feed
        var exportRows = []; // hold all the rows that are generated to be pasted into the sheet
        var XMLFeedURL = FeedURL;
        var feedContent = UrlFetchApp.fetch(XMLFeedURL).getContentText(); // get the full feed content
        var feedItems = XmlService.parse(feedContent).getRootElement().getChild('result').getChild('tasks').getChildren('item'); // get all items in the feed
        var next = XmlService.parse(feedContent).getRootElement().getChildText('next');
      
      
       var nodeArray = ["id", "priority", "title", "deadline","closedDate", "groupId"];
       if(next){
        for (var x = 0; x < feedItems.length; x++) {
            var currentFeedItem = feedItems[x];
            var singleItemArray = [];
          for (var y = 0; y < nodeArray.length; y++) {
              if(nodeArray[y]==="OPPORTUNITY") {
              singleItemArray.push(currentFeedItem.getChildText(nodeArray[y])*1);
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
    var GoogleSheetsPastePage = GoogleSheetsFile.getSheetByName('b24auto');
  GoogleSheetsPastePage.clear();
    GoogleSheetsPastePage.getRange(1, 1, 1, nodeArray.length).setValues([nodeArray]);

  let massive = xmlARR.flat();
        GoogleSheetsPastePage.getRange(2, 1, massive.length, nodeArray.length).setValues(massive);
}
