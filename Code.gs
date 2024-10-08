/**
 * Adds a custom menu to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
      .addItem('Add Attendees to Event', 'addAttendeesToEvent')
      .addToUi();
}

/**
 * Adds attendees from the "Form Responses 1" sheet to a specified Google Calendar event and sends invites.
 */
function addAttendeesToEvent() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Access the "Form Responses 1" sheet for email addresses
    var emailSheet = spreadsheet.getSheetByName("Form Responses 1");
    if (!emailSheet) {
      SpreadsheetApp.getUi().alert('Sheet "Form Responses 1" not found.');
      return;
    }
    
    // Retrieve the range of attendee emails (assuming they start from C2)
    var emailRange = emailSheet.getRange("C2:C" + emailSheet.getLastRow());
    var attendees = emailRange.getValues();
    
    // Access the "CalendarInfo" sheet for Calendar ID and Event ID
    var calendarInfoSheet = spreadsheet.getSheetByName("CalendarInfo");
    if (!calendarInfoSheet) {
      SpreadsheetApp.getUi().alert('Sheet "CalendarInfo" not found.');
      return;
    }
    
    // Retrieve the Calendar ID from cell B2 in "CalendarInfo"
    var calendarId = calendarInfoSheet.getRange("B2").getValue().trim();
    if (!calendarId) {
      SpreadsheetApp.getUi().alert("Calendar ID is missing. Please enter it in cell B2 of the 'CalendarInfo' sheet.");
      return;
    }
    
    Logger.log("Calendar ID: " + calendarId);
    
    // Retrieve the Event ID from cell A2 in "CalendarInfo"
    var eventId = calendarInfoSheet.getRange("A2").getValue().trim();
    if (!eventId) {
      SpreadsheetApp.getUi().alert("Event ID is missing. Please enter it in cell A2 of the 'CalendarInfo' sheet.");
      return;
    }
    
    Logger.log("Original Event ID: " + eventId);
    
    // Clean the Event ID if it contains an underscore (required by Calendar API)
    var cleanedEventId = eventId.includes('_') ? eventId.split('_')[0] : eventId;
    Logger.log("Cleaned Event ID: " + cleanedEventId);
    
    // Retrieve the existing event using the Advanced Calendar API
    var event = Calendar.Events.get(calendarId, cleanedEventId);
    if (!event) {
      SpreadsheetApp.getUi().alert("Event not found. Please check the Event ID in cell A2 of the 'CalendarInfo' sheet.");
      Logger.log("Event not found for ID: " + cleanedEventId);
      return;
    }
    
    Logger.log("Event found: " + event.summary);
    
    // Initialize a list to collect new attendees
    var newAttendees = [];
    var addedCount = 0;
    var skippedCount = 0;
    
    // Create a Set of existing attendee emails to prevent duplicates
    var existingAttendees = new Set();
    if (event.attendees) {
      event.attendees.forEach(function(att) {
        if (att.email) {
          existingAttendees.add(att.email.toLowerCase());
        }
      });
    }
    
    // Iterate through each email and prepare to add as a guest
    for (var i = 0; i < attendees.length; i++) {
      var email = attendees[i][0].trim();
      if (email && validateEmail(email)) { // Validate email format
        if (!existingAttendees.has(email.toLowerCase())) {
          newAttendees.push({ email: email });
          addedCount++;
          Logger.log("Prepared to add guest: " + email);
        } else {
          skippedCount++;
          Logger.log("Skipped existing guest: " + email);
        }
      } else {
        Logger.log("Invalid or empty email at row " + (i + 2) + ": '" + email + "'");
      }
    }
    
    if (newAttendees.length === 0) {
      SpreadsheetApp.getUi().alert("No new valid attendees to add.");
      return;
    }
    
    // Merge existing attendees with new attendees
    event.attendees = event.attendees || [];
    event.attendees = event.attendees.concat(newAttendees);
    
    // Update the event with the new attendees and send updates
    var updatedEvent = Calendar.Events.patch(event, calendarId, cleanedEventId, { sendUpdates: 'all' });
    
    SpreadsheetApp.getUi().alert(
      addedCount + " attendee(s) added successfully!" + 
      (skippedCount > 0 ? "\n" + skippedCount + " attendee(s) were already invited and were skipped." : "")
    );
    
  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
    Logger.log("Error: " + error.toString());
  }
}

/**
 * Validates the email format.
 * @param {string} email The email address to validate.
 * @return {boolean} True if valid, false otherwise.
 */
function validateEmail(email) {
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}
