var calendarId = PropertiesService.getScriptProperties().getProperty("calendarId");
var schedUrl = PropertiesService.getScriptProperties().getProperty("schedUrl");
var HWUrl = PropertiesService.getScriptProperties().getProperty("HWUrl");

var schedHWColumnIndex = 3;
var HWNotFoundMessage = "Details not found: check the HW calendar at ";
function updateCalendar() {
  var chemSchedDoc = DocumentApp.openByUrl(schedUrl);
  var chemHWDoc = DocumentApp.openByUrl(HWUrl);
  var schedTable = chemSchedDoc.getBody().getTables()[0]; // there's only one table in each document
  var HWTable = chemHWDoc.getBody().getTables()[0];
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  var todaysRowIndex = schedTable.getChildIndex(schedTable.findText(dateTableNotation(new Date())).getElement().getParent().getParent().getParentRow());
  
  for (var i = todaysRowIndex; i < schedTable.getNumRows(); i++) { // Start from today and work down the table
    var currentRow = schedTable.getRow(i);
    var currentDate = new Date(currentRow.getCell(0).getText());
    clearEventsOnDay(calendar, currentDate);
    var HWTitles = HWTitlesFromRow(currentRow);
    for (var j = 0; j < HWTitles.length; j++) {
      var HWTitle = HWTitles[j];
      var currentHW = HWTableQueryStringFromTitle(HWTitle);
      var HWContent = HWFromQueryString(currentHW, HWTable);
      if (!HWContent) {
        HWContent = HWNotFoundMessage + HWUrl;
      }
      var ev = calendar.createAllDayEvent(HWTitle, currentDate);
      ev.setDescription(HWContent);
    }
  }
}

function clearEventsOnDay(calendar, date) {
  var eventList = calendar.getEventsForDay(date);
  for (var i = 0; i < eventList.length; i++) {
    eventList[i].deleteEvent();
  }
}

function HWTableQueryStringFromTitle(HWTitle) {
  var wordList = HWTitle.split(' ');
  wordList.shift();
  if (wordList[wordList.length - 1].toUpperCase() === "DUE") wordList.pop(); // pop off the last word if we have any words to spare (this is to catch things like "HW 8.1")
  return wordList.join(' ').trim();
}

function HWTitlesFromRow(schedRow) {
  var cell = schedRow.getCell(schedHWColumnIndex);
  var titles = []
  for (var i = 0; i < cell.getNumChildren(); i++) {
    var listItem = cell.getChild(i);
    var listItemText = stripBulletPoint(cell.getChild(i).getText());
    if (listItemText && listItem.getType() == DocumentApp.ElementType.LIST_ITEM) {
      titles.push(listItemText);
    } else if (listItemText && listItem.getType() == DocumentApp.ElementType.PARAGRAPH) {
      titles[titles.length - 1] += ("\n" + listItemText); // If they forget to make the last few titles in the bulleted format, automatically combine them with the last ListItem
    }
  }
  return titles;
}

function stripBulletPoint(str) {
  var splitString = str.split(' ');
  if (splitString[0] == '●') {
    splitString.shift(); splitString.shift(); // The first two items are the bullet and an extra space
  }
  return splitString.join(' ').trim();
}

function regexEscape(s) {
    return s.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}; 
// TODO TODO TODO TODO: Check if multiple results appear for the query string; if so, return null
// (better safe than sorry)
function HWFromQueryString(HWString, HWTable) {
  try {
    var HWRow = HWTable.findText(regexEscape(HWString)).getElement().getParent().getParent().asTableCell().getParentRow();
    var HWContent = HWRow.getCell(2).getText();
  } catch (e) { // if no match is found for the query string
    var HWContent = null;
  }
  return HWContent;
}

function dateTableNotation(date) {
  return (date.getMonth()+1) + '/' + date.getDate() + '/' + date.getFullYear();
}
