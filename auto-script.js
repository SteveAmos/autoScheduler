//Batch add events to Google Calendar

function eventCreator() {
//Enter Spreadsheet ID
  var mySpreadSheet = SpreadsheetApp.openById("INSERT_SPREADSHEET_ID");

  var testSheet = ("Sheet1");
  var s304 = ("304_WI_23");
  var s306 = ("306_WI_23");
  var s305 = ("305_WI_23");
  var s300 = ("300_WI_23");

//Select a sheet variable from above and enter below to switch source sheets:
  var sheet = mySpreadSheet.getSheetByName(s305);

  var last_row = sheet.getLastRow();
  var data = sheet.getRange("A2:F"+ last_row).getValues();

//Enter Calendar ID
  var cal_304 = CalendarApp.getCalendarById("INSERT_CALENDAR_ID_@group.calendar.google.com");
  var cal_305 = CalendarApp.getCalendarById("INSERT_CALENDAR_ID_@group.calendar.google.com");
  var cal_306 = CalendarApp.getCalendarById("INSERT_CALENDAR_ID_@group.calendar.google.com");
  var cal_300 = CalendarApp.getCalendarById("INSERT_CALENDAR_ID_@group.calendar.google.com");

//Select a calendar ID variable from above and enter below to switch calendars:
  var cal = (cal_305);
  
//Logger.log(cal.getName());

for(var i = 0; i < data.length; i++){

//
var event = cal.createEventSeries(data[i][0],
    new Date(data[i][1]),
    new Date(data[i][2]),
    CalendarApp.newRecurrence().addWeeklyRule().times(data[i][3]),
    {location: data[i][4], description:data[i][5] });
}

} //...end of eventCreator