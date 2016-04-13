//Script to take a class schedule, add it to a calendar and create a form for attendees to register

//FUNCTION to add a custom menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Getting started", functionName: "getStart"});
  menuEntries.push({name: "Add calendar data", functionName: "addData"});
  menuEntries.push({name: "Create Calendar", functionName: "createCalendar"});
  menuEntries.push({name: "Reset Calendar", functionName: "resetCalendar"});
  sheet.addMenu("SAKG Classes", menuEntries);  
}

//FUNCTION show modal window with Getting Started instructions
function getStart() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Welcome to the SAKG Roster Management calendar');
}


//FUNCTION to list current calendars

function listCalendars() {
  var calendars, pageToken;
  do {
    calendars = Calendar.CalendarList.list({
      maxResults: 100,
      pageToken: pageToken
    });
    if (calendars.items && calendars.items.length > 0) {
      for (var i = 0; i < calendars.items.length; i++) {
        var calendar = calendars.items[i];
        Logger.log('%s (ID: %s)', calendar.summary, calendar.id);
      }
    } else {
      Logger.log('No calendars found.');
    }
    pageToken = calendars.nextPageToken;
  } while (pageToken);
}


//FUNCTION to populate Raw Data with events list
function addData() {

  //create Raw Data sheet first
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var rawData = activeSpreadsheet.getSheetByName("Raw Data"); 

  //push formulas into Raw Data sheet
  var sheet = activeSpreadsheet.getSheetByName("SAKG Classes");
  var data = sheet.getDataRange().getValues();
  var lastRow = sheet.getLastRow();
  
  var formulas = [["=CONCATENATE('SAKG Classes'!A2,\" - \",'SAKG Classes'!B2,\" on \",Text('SAKG Classes'!C2, \"dd-MMMM\"),\" at \",Text('SAKG Classes'!D2, \"h:mm AM/PM\"))","='SAKG Classes'!A2","='SAKG Classes'!B2","='SAKG Classes'!C2+'SAKG Classes'!D2","='SAKG Classes'!C2+'SAKG Classes'!E2","=CONCATENATE(\"Max \",'SAKG Classes'!F2,\" - \",'SAKG Classes'!G2)"]];
  var cell = rawData.getRange("A2:F2");
  cell.setFormulas(formulas); 
  
  var rangeToCopy = rawData.getRange(2, 1, 1, rawData.getMaxColumns());
  for (n=2; n<=data.length; n++){
    var count = n;
  rangeToCopy.copyTo(rawData.getRange(count, 1));  
  }
}  


//FUNCTION to push new events to calendar
function createCalendar() {
  
  //export volunteers to Raw Data sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("volunteers");
  var data = sheet.getDataRange().getValues();
  var destination = ss.getSheetByName("Raw Data")
  var rawData = destination.getDataRange().getValues();

  //get list of sessions
  var sessionList = [];
  for(n in data){
    var session = data[n][0];
    var duplicate = false;
    for (j in sessionList){
      if(session == sessionList[j]){
        duplicate = true;
      }
    }
    if(!duplicate){
      sessionList.push(session);
    }
  }

  //add volunteers for each session
  for (n in sessionList) {
    var volunteers = [];
    for (i in data){
      if (sessionList[n] == data[i][0]){
        volunteers.push(data[i][1]);
      }
    }
      for (j in rawData) {
        Logger.log(sessionList[n]);
        Logger.log(rawData[j][0]);
        if (sessionList[n] == rawData[j][0]){
          var sessionVolunteers = volunteers.join(', ');
          h = Number(j) + 1;
          destination.getRange(h,8).setValue(sessionVolunteers);
      }
    }
  }

  
  //spreadsheet variables
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Raw Data");
  var range = sheet.getDataRange();
  var values = range.getValues();    
  
  //calendar variables
  var calendar = CalendarApp.getCalendarById('q3kh2950cfha2i3n503qv8999o@group.calendar.google.com'); //input your calendar ID 

  for (var i = 1; i < values.length; i++) {     
    if (values[i][8] != 'Added') {                //to avoid duplicates, check if it's been entered before
      var eventTitle = 'SAKG: ' + values[i][1] + ' - '+ values[i][2];
      var eventDescription = {description: values[i][7]};
      var start = values[i][3];
      var end = values[i][4];
  
      var event = calendar.createEvent(eventTitle, start, end, eventDescription);
                
      //get ID
      var eventId = event.getId();
        
      //mark as entered, enter ID
      sheet.getRange(i+1,9).setValue('Added');
      sheet.getRange(i+1,10).setValue(eventId);
      }
   }

  //Add modal window with instructions
  var html = HtmlService.createHtmlOutputFromFile('Create Calendar')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'SAKG Classes added to your Google calendar!');
}


//FUNCTION to remove all events from calendar
function resetCalendar() {
  
  var fromDate = new Date(2015,0,1,0,0,0);
  var toDate = new Date(2018,0,1,0,0,0);
  var calendar = CalendarApp.getCalendarById('q3kh2950cfha2i3n503qv8999o@group.calendar.google.com'); //input your calendar ID 

// delete from Jan 1 to end of Jan 1, 2018

  var events = calendar.getEvents(fromDate, toDate);
  for(var i=0; i<events.length;i++){
    var ev = events[i];
    Logger.log(ev.getTitle()); // show event name in log
    ev.deleteEvent();
  }
}