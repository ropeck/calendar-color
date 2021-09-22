/*  go/ropeck-auto-calendar-colors
 *
 * copied from Erik's note in b/25695571 erikdejong@  go/erdj-calendar-colours
 *
 * Inspired by https://rickpastoor.com/2019/05/30/google-calendar-color-coder.html
 *
 *
 * Installation Instruction:
 * 1. Make sure the Calendar API is enabled under
 *      Resources -> Advanced Google Services
 * 2. Run the 'installAndRun' function
 */

// https://stackoverflow.com/questions/14221264/open-google-docs-spreadsheet-by-name
function getSpreadsheetByName(filename) {
  var files = DriveApp.getFilesByName(filename);

  for(var i in files)
  {
    if(files[i].getName() == filename)
    {
      // open - undocumented function
      return SpreadsheetApp.open(files[i]);
      // openById - documented but more verbose
      // return SpreadsheetApp.openById(files[i].getId());
    }
  }
  return null;
}

function testsheet() {
  var sheet = getSpreadsheetByName("interviews");
}

function create_event(calendarId, subject, start, description, color = CalendarApp.EventColor.GREEN) {
  var cal = CalendarApp.getCalendarById(calendarId);
  var feed = cal.createEvent(subject, new Date(start.getTime()), new Date(start.getTime() + 60*60*1000));
  feed.setColor(color);
  feed.setDescription(description);
  return feed;
}

function handle_interview(calendarId, event) {
  Logger.log('handle_interview: ' + event.summary);
 // # create prep and post calendar entries and interview doc
    // Calendar.Events.insert(calendarId, prep);
    //     var date = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd"); // "yyyy-MM-dd'T'HH:mm:ss'Z'"
  var date = Utilities.formatDate(new Date(event.start.dateTime), "PDT", "yyyyMMdd");
  var name = event.summary.match(/Interview with (.*)/);
  if (name) {
    name = name[1];
  }
  var docname = date + " interview: " + name;
  // var folder = DriveApp.getFoldersByName("interviews").next();
  if (DriveApp.getFilesByName(docname).hasNext() === false) {
    // create the doc
    var blob = HtmlService.createHtmlOutput(event.description).getBlob()
    var newFileId = Drive.Files.insert({title: docname+"-html"}, blob, {convert: true}).id;
    doc = DocumentApp.create(docname);
    var body = doc.getBody();
    if (body) {
      body.appendParagraph(event.start.dateTime);
      body.appendTable([[event.description]]);
     // body.appendTable([[blob]]);
    } 
    
    // create the pre and post calendar events
    var interview = new Date(event.start.dateTime);
    var doclink = "<a href='"+doc.getUrl()+"'>Interview Notes</a>";
    // prep interview time should be on the Friday before if the interview is on a Monday
    var prepdays = 1;
    if (interview.getDay() == 1) {
      prepdays = 3;
    }

    create_event(calendarId, "prep interview: " + name, new Date(interview.getTime() - prepdays * 24 * 60 * 60 * 1000), doclink);
    create_event(calendarId, "feedback interview: " + name, new Date(interview.getTime() + 60*60*1000), doclink);
  
  //   // add a row to the interview tracking spreadsheet
  //   var ss = getSpreadsheetByName("interviews");
  //   var sheet = ss.getSheets()[0];
  
  // // Appends a new row with 3 columns to the bottom of the
  // // spreadsheet containing the values in the array
  //   sheet.appendRow(["a man", "a plan", "panama"]);
  }
}

function duplicateEvent(calendarId, event) {
  var oneWeekAgo = new Date();
  oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
  if (new Date(event.start.dateTime) > oneWeekAgo) {
    return;
  }
  var cal = CalendarApp.getCalendarById(calendarId);
  var events = cal.getEvents(new Date(event.start.dateTime), new Date(event.end.dateTime));
  var count = 0;
  events.forEach(function(x) {
    if (x.getSummary() == event.getSummary()) {
      count = count + 1;
    }
  });
  Logger.log("count " + count + " " + event.summary);
  if (count != 1) {
    return;
  }
  var eventCopy={};
  eventCopy.summary = event.summary;
  eventCopy.description = event.description!=undefined?event.description:event.summary;
  eventCopy.description = "(copy) " + eventCopy.description;
  eventCopy.location = event.location;
  eventCopy.start = event.start;
  eventCopy.end = event.end;
  eventCopy.recurrence = event.recurrence;
  eventCopy.attendees = event.attendees;
  eventCopy.colorId = event.colorId;
  copycal = Calendar.Events.insert(eventCopy,calendarId);//duplique ds le même agenda
  Logger.log("duplicated " + event.summary);
}

function decorateEvent(calendarId, event, reg, color, duplicate = true, extra = false) {
  if (event.summary.match(reg)) {
    if (event.colorId != color) {
      Calendar.Events.patch({ colorId: color }, calendarId, event.id);
      Logger.log('color set: ' + event.colorId + ' ' + event.summary);
    }
    if (extra) {
      extra(calendarId, event);
    }
    var undef_desc = typeof(event.description) != "undefined";
    if (duplicate && (! undef_desc || ! event.description.match(/(copy)/))) {
      duplicateEvent(calendarId, event);
      var desc = event.description!=undefined?event.description:'';
      Logger.log(reg.source + ' ' + event.start + ' ' + desc.substr(0,20));
    }
  }
}

function processCalendarEvent_(calendarId, event, history) {
  history[event.summary] = event
  var msg = '';
  if (event.summary) {
    msg = event.summary.substr(0,120);
  }
  Logger.log('process: ' + ' ' + event.start + ' ' + msg);
  // decorateEvent(calendarId, event, / cloud-support-/, CalendarApp.EventColor.PALE_RED);
  decorateEvent(calendarId, event, /Oncall/, CalendarApp.EventColor.PALE_RED);
  decorateEvent(calendarId, event, /Cases and Consults Queue/, CalendarApp.EventColor.PALE_BLUE);
  decorateEvent(calendarId, event, /Onsite.*Interview/, CalendarApp.EventColor.PALE_GREEN, false, handle_interview);
}

/*
 * Install the trigger and perform a complete sync on the primary calendar.
 */
function installAndRun() {
  installTrigger();
  
  var calendarId = 'primary'; //CalendarApp.getDefaultCalendar().getId();
  updateSyncedEvents_(calendarId, true);
}

/*
 * Install the trigger for new and modified events.
 */
function installTrigger() {
  // First delete all triggers in the current project.
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  
  // Then install the calendar update event trigger
  ScriptApp.newTrigger('colourModifiedEvents')
    .forUserCalendar(Session.getActiveUser().getEmail())
    .onEventUpdated()
    .create()
}

/*
 * Process the modified events on the primary calendar.
 */
function colourModifiedEvents() {
  var calendarId = 'primary'; //CalendarApp.getDefaultCalendar().getId();
  updateSyncedEvents_(calendarId, false);
}

function colourAllEvents() {
  var calendarId = 'primary'; //CalendarApp.getDefaultCalendar().getId();
  updateSyncedEvents_(calendarId, true);
}

/**
 * Symchronize new and modified events for a calendar.
 * Modified from:
 *  https://developers.google.com/apps-script/advanced/calendar#synchronizing_events
 * @param {string} calendarId The identifier of the calendar to sync.
 * @param {boolean} fullSync Synchronize all events as opposed to the modified ones since the last sync.
 */
function updateSyncedEvents_(calendarId, fullSync) {
  Logger.log("Calendar: " + calendarId);
  var properties = PropertiesService.getUserProperties();
  var options = {
    maxResults: 100,
    showDeleted: true
  };
  var syncTokenKey = 'syncToken:' + calendarId;
  var syncToken = properties.getProperty(syncTokenKey);
  Logger.log('sync token ' + syncTokenKey  + ' ' + syncToken + ' full sync: ' + fullSync);
  if (syncToken && !fullSync) {
    options.syncToken = syncToken;
  } else {
    // Sync events up to thirty days in the past.
    options.timeMin = getRelativeDate_(-30, 0).toISOString();
  }

  // Retrieve events one page at a time.
  var events;
  var pageToken;
  var history = {};
  do {
    try {
      options.pageToken = pageToken;
      Logger.log('options: ' + JSON.stringify(options));
      events = Calendar.Events.list(calendarId, options);
      Logger.log('saving sync token' + syncTokenKey  + ' ' + events.nextSyncToken);
      if (events.nextSyncToken) {
        properties.setProperty(syncTokenKey, events.nextSyncToken);
      }
    } catch (e) {
      Logger.log('error:' + e.message);
      // Check to see if the sync token was invalidated by the server;
      // if so, perform a full sync instead.
      if (e.message === 'Sync token is no longer valid, a full sync is required.') {
        properties.deleteProperty('syncToken');
        updateSyncedEvents_(calendarId, true);
        return;
      } else {
        throw new Error(e.message);
      }
    }

    if (events.items && events.items.length > 0) {
      for (var i = 0; i < events.items.length; i++) {
        var event = events.items[i];
        var desc = '';
        if (event.summary) {
          desc = event.summary.substr(0,1024);
        }
        Logger.log('event: ' + desc + ' : ' + JSON.stringify(event).substr(0,500));
        if (event.status === 'cancelled') {
          continue;
        }
        processCalendarEvent_(calendarId, event, history);
      }
    } else {
      Logger.log('No events found.');
    }

    pageToken = events.nextPageToken;
  } while (pageToken);
  Logger.log('saving sync token' + syncTokenKey  + ' ' + events.nextSyncToken)
  properties.setProperty(syncTokenKey, events.nextSyncToken);
}

/**
 * Helper function to get a new Date object relative to the current date.
 * @param {number} daysOffset The number of days in the future for the new date.
 * @param {number} hour The hour of the day for the new date, in the time zone
 *     of the script.
 * @return {Date} The new date.
 */
function getRelativeDate_(daysOffset, hour) {
  var date = new Date();
  date.setDate(date.getDate() + daysOffset);
  date.setHours(hour);
  date.setMinutes(0);
  date.setSeconds(0);
  date.setMilliseconds(0);
  return date;
}


/*  https://engineering.continuity.net/test-for-google-apps-script/ */

/* unit testing */

// Optional for easier use.

function doGet() {
var QUnit = QUnitGS2.QUnit;
   QUnitGS2.init(); // Initializes the library.

   /*
   * Add your test functions here.
   */


  QUnit.module('add', function() {

   QUnit.test("another example", function( assert ) {
    assert.equal(10, 10);
   });

   QUnit.test( "Arrays basics", function( assert ) {
    assert.equal( QUnit.equiv( ['one'], []), true );
   });
  });
   QUnit.start(); // Starts running tests, notice QUnit vs QUnitGS2.
   return QUnitGS2.getHtml();
}

function getResultsFromServer() {
   return QUnitGS2.getResultsFromServer();
}

