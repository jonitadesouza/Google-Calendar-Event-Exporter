function exportAllCalendarEventsWithParticipants() {
  var calendarId = 'primary'; // Change this to a specific calendar ID if needed
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  var events = calendar.getEvents(new Date(0), new Date()); // Get all past and future events
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Clear previous data
  sheet.clear();
  
  // Set headers
  sheet.appendRow(["Event Title", "Start Time", "End Time", "Description", "Location", "Participants"]);

  // Add events to the sheet
  events.forEach(event => {
    var attendeesList = "None"; // Default value if no attendees

    if (event.getGuestList) { // Ensure the method exists
      var attendees = event.getGuestList(); // Get attendees
      if (attendees.length > 0) {
        attendeesList = attendees.map(guest => guest.getEmail()).join(", "); // Join emails with commas
      }
    }
    
    sheet.appendRow([
      event.getTitle(),
      event.getStartTime(),
      event.getEndTime(),
      event.getDescription(),
      event.getLocation(),
      attendeesList
    ]);
  });

  SpreadsheetApp.getUi().alert("All calendar events with participants exported successfully!");
}
