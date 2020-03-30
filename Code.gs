/*
*=========================================
*       INSTALLATION INSTRUCTIONS
*=========================================
*
* 1) Click in the menu "File" > "Make a copy..." and make a copy to your Google App Script environment
* 2) Click in the menu "Run" > "Run function" > "run" and authorize the program
* 3) Press ctrl + Enter and note your "contact calendar ID"
* 4) Changes lines 19-26 to be the settings that you want to use
* 5) Click in the menu "Run" > "Run function" > "install" tu run the programm periodically
*
* **To stop the Script from running click in the menu "Run" > "Run function" > "uninstall"
*/

/*
*=========================================
*               SETTINGS
*=========================================
*/
//var calendarId = '';
var calendarId = 'addressbook#contacts@group.v.calendar.google.com';    // Contact calendar ID. Can be found when running the "run" script with calendarId = ''
var newCalendarName = 'Anniversaires';                                  // Calendar Name
var calendarColor = "#3F51B5";                                          // Calendar Color
var maxEventToCopy = 10;                                                // What interval (minutes) to run this script on to check for new events
var eventColor = CalendarApp.EventColor.PALE_BLUE;                      // Event color
var eventTimeNotification = 12;                                         // Hour of the event notification
var howFrequent = 720;                                                  // What interval (minutes) to run this script on to check for new events

function install(){
  unInstall();

  if (howFrequent < 1){
    throw "[ERROR] \"howFrequent\" must be greater than 0.";
  }
  else{
    ScriptApp.newTrigger("run").timeBased().after(1000).create();//Start the sync routine
  }
}

function unInstall(){
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "run") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function run() {
  if (calendarId === '') listCalendars();
  else {
    clearCalendar(newCalendarName);
    var newCalendar = createCalendar(newCalendarName);
    var events = getNextEvents(calendarId);
    addEvents(newCalendar, events);
  }
}

function listCalendars() {
  var calendars;
  var pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        Logger.log('%s (ID: %s)', calendar.summary, calendar.id, calendar.name);
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}

function createCalendar(name) {
  var id = "";
  var calendar;

  // Gets the calendar
  var calendars = CalendarApp.getCalendarsByName(name);

  if (calendars.length == 0) {
    // Creates a new calendar
    calendar = CalendarApp.createCalendar(name, {hidden : false, selected: true, color: calendarColor});
    Logger.log('Created the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
  }
  else {
    calendar = calendars[0];
    Logger.log('Found the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
  }
  
  return calendar;
}

function deleteCalendar(name) {
  var id = "";
  var calendar;

  // Gets the calendar
  var calendars = CalendarApp.getCalendarsByName(name);

  if (calendars.length > 0) {
    calendar = calendars[0];
    Logger.log('Deleted the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
    calendar.deleteCalendar();
    
    return true;
  }
  
  return false;
}

function clearCalendar(name) {
  var id = "";
  var calendar;

  // Gets the calendar
  var calendars = CalendarApp.getCalendarsByName(name);

  if (calendars.length > 0) {
    calendar = calendars[0];
    var date = new Date();
    var firstDay = new Date(date.getFullYear(), date.getMonth() - 1, 1);
    var lastDay = new Date(date.getFullYear(), date.getMonth() + 6, 1);
    
    var events = calendar.getEvents(firstDay, lastDay);
    for (i = 0; i < events.length; i++) {
      var event = events[i];
      event.deleteEvent();
    }
    Logger.log('Cleared the calendar "%s", with the ID "%s".', calendar.getName(), calendar.getId());
    
    return true;
  }
  
  return false;
}

function getNextEvents(calendarId) {
  var optionalArgs = {
    timeMin: (new Date()).toISOString(),
    showDeleted: false,
    singleEvents: true,
    maxResults: maxEventToCopy,
    orderBy: 'startTime'
  };
  var response = Calendar.Events.list(calendarId, optionalArgs);
  var events = response.items;
  if (events.length == 0) Logger.log('No upcoming events found.');
  
  return events;
}

function addEvents(calendar, events) {
  for (i = 0; i < events.length; i++) {
    var event = events[i];
    var dateTime = new Date(event.start.date);
    dateTime.setHours(eventTimeNotification);
    
    Logger.log('%s (%s)', event.summary, dateTime);
    event = calendar.createEvent(event.summary, dateTime, dateTime);
    event.setColor(eventColor);
    event.addPopupReminder(5);
  }
}
