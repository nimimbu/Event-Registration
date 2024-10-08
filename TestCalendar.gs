function testCalendarAccess() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var calendarId = sheet.getRange("B2").getValue().trim();
  
  if (!calendarId) {
    SpreadsheetApp.getUi().alert("Calendar ID is missing. Please enter it in cell B2.");
    return;
  }
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  if (!calendar) {
    SpreadsheetApp.getUi().alert("Calendar not found. Please check the Calendar ID in cell B2.");
    return;
  }
  
  SpreadsheetApp.getUi().alert("Calendar accessed successfully: " + calendar.getName());
}
