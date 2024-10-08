/**
 * Lists upcoming events in the specified calendar and writes their titles and IDs to the sheet.
 */
function listUpcomingEvents() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Get the calendar ID from the sheet (e.g., cell B2)
    var calendarId = sheet.getRange("B2").getValue().trim();
    
    if (!calendarId) {
      SpreadsheetApp.getUi().alert("Calendar ID is missing. Please enter it in cell B2.");
      return;
    }
    
    // Get the calendar by ID
    var calendar = CalendarApp.getCalendarById(calendarId);
    
    if (!calendar) {
      SpreadsheetApp.getUi().alert("Calendar not found. Please check the Calendar ID in cell B2.");
      return;
    }
    
    // Define the time range for upcoming events (e.g., next 30 days)
    var now = new Date();
    var future = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000); // 30 days from now
    
    var events = calendar.getEvents(now, future);
    
    if (events.length === 0) {
      SpreadsheetApp.getUi().alert("No upcoming events found in the next 30 days.");
      return;
    }
    
    // Prepare data to write to the sheet
    var output = [["Event Title", "Event ID"]];
    for (var i = 0; i < events.length; i++) {
      var event = events[i];
      output.push([event.getTitle(), event.getId()]);
    }
    
    // Clear previous output (optional)
    sheet.getRange("A4:B").clearContent();
    
    // Write the event titles and IDs starting from cell A4
    sheet.getRange(4, 1, output.length, output[0].length).setValues(output);
    
    SpreadsheetApp.getUi().alert("Listed " + events.length + " upcoming events in cells A4:B.");
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
    Logger.log("Error: " + error.toString());
  }
}
