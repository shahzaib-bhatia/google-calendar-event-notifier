const spreadsheetParent = SpreadsheetApp.getActiveSpreadsheet();
SpreadsheetApp.setActiveSheet(spreadsheetParent.getSheets()[1]);
const spreadsheet = SpreadsheetApp.getActiveSheet();

const zoomRoomID = "..."
const zoomReminder = "..."

const StaffCalendarID = "..."
const DirectorCalendarID = "..."

const StaffCalShareURL = "..."
const DirectorCalShareURL = "..."

const StaffChatURL = "..."
const DirectorChatURL = "..."

const StaffEventCal = CalendarApp.getCalendarById(StaffCalendarID);
const DirectorEventCal = CalendarApp.getCalendarById(DirectorCalendarID);
const meetings = spreadsheet.getRange("A2:M").getValues();

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync to Calendar')
      .addItem('Sync', 'updateCalendar')
      .addToUi();
}

function updateCalendar() {
  clearCalendar(StaffEventCal);
  clearCalendar(DirectorEventCal);
  for (x=0; x<meetings.length; x++) {
    var meeting = meetings[x];
      if (meeting [0] && meeting[5] && meeting[6]) { 
        var title = meeting[0];
        var start = meeting[5];
        var end = meeting[6];
        var location = meeting[8];
        var description = meeting[9];
        var audience = meeting[10];
        var zoom = meeting[11];
        var notify = meeting[12];

        Logger.log("Adding " + (x + 1) + " / " + meetings.length + " - " + title + " Start: " + start + " End: " + end + " Location: " + location + " Description: " + description);

      if ( audience == "Director") {
          addEvent(DirectorEventCal,title,start,end,location,description,zoom,notify);
        } else {
          addEvent(StaffEventCal,title,start,end,location,description,zoom,notify);
      }
    }
  } 
}

function addEvent(calendar,title,start,end,location,description,zoom,notify) {
  if ( zoom == "Yes" ) { description = description + zoomReminder };
  var options = {location: location, description: description};
  var event = calendar.createEvent(title,start,end,options);
  if ( zoom == "Yes" ) { event.addGuest(zoomRoomID) };
  if ( notify == "Yes" ) { event.addPopupReminder(30) };
  Logger.log("OK! Created: " + event.getId() );
  Utilities.sleep(250); // Api limit
}

function clearCalendar(eventCal) {
  const start = new Date('January 1 1970');
  const end = new Date('December 31 2077');

  const events = eventCal.getEvents(start, end)

  for (x=0; x<events.length; x++) {
    var event = events[x];
    Logger.log("Removing event " + ( x + 1 ) + " / " + events.length + " " + event.getId() )
    event.deleteEvent();
    Utilities.sleep(250) // Api limit
  }

}

/* 
function dailyChatReminder() {
  var message = [];
  message = eventJSON(1);
  if (message.length > 0 ) {
    sendMessage(
      "Meetings/Events today!", 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/today/default/24px.svg",
      message);
  } else {
    
  }
}
*/

function dailyChatReminderFiltered() {
  const windowSize=1
  messageStaff = eventJSONFiltered(windowSize,["Staff","Public"]);
  messageDirector = eventJSONFiltered(windowSize,["Director"]);
  
  if (messageStaff.length > 0 ) {
    sendMessageTo(
      "Meetings/Events today!", 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/today/default/24px.svg",
      messageStaff,StaffChatURL,StaffCalShareURL);
  } else {
    
  }
  if (messageDirector.length > 0 ) {
    sendMessageTo(
      "Director Meetings/Events today!", 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/today/default/24px.svg",
      messageDirector,DirectorChatURL,DirectorCalShareURL);
  } else {
    
  }
}

/*
function weeklyChatReminder() {
  const windowSize=28
  const today=new Date();
  const later=new Date(today.getFullYear(), today.getMonth(), today.getDate() + windowSize);
  var message = [];
  const titleOptions = {
    month: "short",
    day: "2-digit",
  };
  message = eventJSON(windowSize);
  if (message.length > 0 ) {
    // everything is fine
  } else {
    message.push({
          header: "<b>No events in selected range</b>",
          // collapsible: true,
          widgets: [
            { textParagraph: {
              maxLines: 3,
              text: 
                "Either the calendar is broken or it's a dry month."
              }
            }
          ]
      });
  }
  sendMessage(
      "Upcoming Meetings/Events" + " — " + new Intl.DateTimeFormat("en-US", titleOptions).format(today) + " to " + new Intl.DateTimeFormat("en-US", titleOptions).format(later), 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/calendar_month/default/24px.svg",
      message);
}
*/

function weeklyChatReminderFiltered() {
  const windowSize=28
  const today=new Date();
  const later=new Date(today.getFullYear(), today.getMonth(), today.getDate() + windowSize);
  var messageStaff = [];
  var messageDirector = [];
  const titleOptions = {
    month: "short",
    day: "2-digit",
  };
  messageStaff = eventJSONFiltered(windowSize,["Staff","Public"]);
  messageDirector = eventJSONFiltered(windowSize,["Director"]);

  if (messageStaff.length > 0 ) {
    // everything is fine
  } else {
    messageStaff.push({
          header: "<b>No events in selected range</b>",
          // collapsible: true,
          widgets: [
            { textParagraph: {
              maxLines: 3,
              text: 
                "Either the calendar is broken or it's a dry month."
              }
            }
          ]
      });
  }
  sendMessageTo(
      "Upcoming Meetings/Events" + " — " + new Intl.DateTimeFormat("en-US", titleOptions).format(today) + " to " + new Intl.DateTimeFormat("en-US", titleOptions).format(later), 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/calendar_month/default/24px.svg",
      messageStaff,StaffChatURL,StaffCalShareURL);

  if (messageDirector.length > 0 ) {
    // everything is fine
  } else {
    messageDirector.push({
          header: "<b>No events in selected range</b>",
          // collapsible: true,
          widgets: [
            { textParagraph: {
              maxLines: 3,
              text: 
                "Either the calendar is broken or it's a dry month."
              }
            }
          ]
      });
  }
  sendMessageTo(
      "Upcoming Director Meetings/Events" + " — " + new Intl.DateTimeFormat("en-US", titleOptions).format(today) + " to " + new Intl.DateTimeFormat("en-US", titleOptions).format(later), 
      "https://fonts.gstatic.com/s/i/short-term/release/googlesymbols/calendar_month/default/24px.svg",
      messageDirector,DirectorChatURL,DirectorCalShareURL);

}

/*
function eventJSON(windowSize) {
  const today=new Date();
  const later=new Date(today.getFullYear(), today.getMonth(), today.getDate() + windowSize);
  const message = [];
  const titleOptions = {
    weekday: "long",
    month: "short",
    day: "2-digit",
  };
  const bodyOptions = {
    hour: "numeric",
    minute: "numeric",
    timeZone: "America/Chicago",
    // timeZoneName: "short",
  };

  for (x=0; x<meetings.length; x++) {    
    var meeting = meetings[x];
      if (meeting [0] && meeting[5] && meeting[6]) { 
        var title = meeting[0];
        var start = meeting[5];
        var end = meeting[6];
        var location = meeting[8];
        var description = meeting[9];
        //var audience = meeting[10];
        // var zoom = meeting[11];
        // var notify = meeting[12];

        if ( today < start && start < later ) {
          message.push({
            header: "<b>" + title + "</b>" + " - " + new Intl.DateTimeFormat("en-US", titleOptions).format(start),
            // collapsible: true,
            widgets: [
              { textParagraph: {
                maxLines: 3,
                text: 
                  "<b>Start:</b> " + new Intl.DateTimeFormat("en-US", bodyOptions).format(start) + " " +
                  "<b>End:</b> " + new Intl.DateTimeFormat("en-US", bodyOptions).format(end) + "<br>" +
                  "<b>Location:</b> " + location + "<br>" +
                  "<b>Description:</b> " + description + "<br>"
                }
              }
            ]
        });
        Logger.log(JSON.stringify(message));
      }
    }
  }
  return message;
}
*/

function eventJSONFiltered(windowSize,filter) {
  const today=new Date();
  const later=new Date(today.getFullYear(), today.getMonth(), today.getDate() + windowSize);
  const message = [];
  const titleOptions = {
    weekday: "long",
    month: "short",
    day: "2-digit",
  };
  const bodyOptions = {
    hour: "numeric",
    minute: "numeric",
    timeZone: "America/Chicago",
    // timeZoneName: "short",
  };

  for (x=0; x<meetings.length; x++) {    
    var meeting = meetings[x];
      if (meeting [0] && meeting[5] && meeting[6]) { 
        var title = meeting[0];
        var start = meeting[5];
        var end = meeting[6];
        var location = meeting[8];
        var description = meeting[9];
        var audience = meeting[10];
        // var zoom = meeting[11];
        // var notify = meeting[12];

        if ( today < start && start < later && filter.includes(audience) ) {
          message.push({
            header: "<b>" + title + "</b>" + " - " + new Intl.DateTimeFormat("en-US", titleOptions).format(start),
            // collapsible: true,
            widgets: [
              { textParagraph: {
                maxLines: 3,
                text: 
                  "<b>Start:</b> " + new Intl.DateTimeFormat("en-US", bodyOptions).format(start) + " " +
                  "<b>End:</b> " + new Intl.DateTimeFormat("en-US", bodyOptions).format(end) + "<br>" +
                  "<b>Location:</b> " + location + "<br>" +
                  "<b>Description:</b> " + description + "<br>"
                }
              }
            ]
        });
        Logger.log(JSON.stringify(message));
      }
    }
  }
  return message;
}

/*
function sendMessage(title, image, inner_payload) {

  const payload = {
  "cardsV2":[
    {
      "card":{
        "header":{
          "title" :title,
          "imageUrl": image
        },
        "sections": inner_payload
      }
    }
  ],
  "accessoryWidgets":[
    {
      "buttonList":{
        "buttons":[
          {
            "text":"Full Schedule",
            "icon":{
              "materialIcon":{
                "name":"link"
              }
            },
            "onClick":{
              "openLink":{
                "url":"..."
              }
            }
          },{
            "text":"Add to Calendar",
            "icon":{
              "materialIcon":{
                "name":"link"
              }
            },
            "onClick":{
              "openLink":{
                "url":"..."
              }
            }
          },
        ]
      }
    }
  ]
};

    const options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
    };
    response=UrlFetchApp.fetch(StaffChatURL, options);
    console.log(response.toString());
}
*/

function sendMessageTo(title, image, inner_payload, target, calendarURL) {

  const payload = {
  "cardsV2":[
    {
      "card":{
        "header":{
          "title" :title,
          "imageUrl": image
        },
        "sections": inner_payload
      }
    }
  ],
  "accessoryWidgets":[
    {
      "buttonList":{
        "buttons":[
          {
            "text":"Full Schedule",
            "icon":{
              "materialIcon":{
                "name":"link"
              }
            },
            "onClick":{
              "openLink":{
                "url":"..."
              }
            }
          },{
            "text":"Add to Calendar",
            "icon":{
              "materialIcon":{
                "name":"link"
              }
            },
            "onClick":{
              "openLink":{
                "url": calendarURL
              }
            }
          },
        ]
      }
    }
  ]
};

    const options = {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
    };
    response=UrlFetchApp.fetch(target, options);
    console.log(response.toString());
}
