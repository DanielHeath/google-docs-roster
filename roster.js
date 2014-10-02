var FUS1=new Date().toString().substr(25,6)+':00';
function onOpen() {
  var menu = [
    // {name: 'Replace roster from calendar', functionName: 'setUpRoster_'},
    {name: 'Import Calendar', functionName: 'Cal_to_sheet'},
    {name: 'Create Form', functionName: 'CreateForm '}
  ];
  SpreadsheetApp.getActive().addMenu('Rostering', menu);

  ss.addMenu("Calendar Uitilities", menuEntries);

 var ui = DocumentApp.getUi();
 var response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

}

/**
 * Creates a Google Form that allows respondents to select which roster
 * sessions they can make it to.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {String[][]} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {

  // Create the form and add a multiple-choice question for each timeslot.
  var form = FormApp.create('Conference Form');

  var item = form.addListItem();
  item.setTitle('Choose a time you are available')
     .setChoices([
         item.createChoice('Cats'),
         item.createChoice('Dogs')
     ]);


  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Mobile number').setRequired(true);
  for (var day in schedule) {
    var header = form.addSectionHeaderItem().setTitle('Sessions for ' + day);
    for (var time in schedule[day]) {
      var item = form.addMultipleChoiceItem().setTitle(time + ' ' + day)
          .setChoiceValues(schedule[day][time]);
    }
  }
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var user = {
    name: e.namedValues['Name'][0],
    email: e.namedValues['Email'][0],
    email: e.namedValues['Timeslot'][0]
  };

  // Grab the session data again so that we can match it to the user's choices.
  var response = [];
  var values = SpreadsheetApp.getActive().getSheetByName('Conference Setup')
     .getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var title = session[0];
    var day = session[1].toLocaleDateString();
    var time = session[2].toLocaleTimeString();
    var timeslot = time + ' ' + day;

    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[timeslot] && e.namedValues[timeslot] == title) {
      response.push(session);
    }
  }
  sendInvites_(user, response);
  sendDoc_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendInvites_(user, response) {
  var id = ScriptProperties.getProperty('calId');
  var cal = CalendarApp.getCalendarById(id);
  for (var i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][5]).addGuest(user.email);
  }
}

/**
 * Create and share a personalized Google Doc that shows the user's itinerary.
 *
 * @param {Object} user An object that contains the user's name and email.
 * @param {String[][]} response An array of data for the user's session choices.
 */
function sendDoc_(user, response) {
  var doc = DocumentApp.create('Conference Itinerary for ' + user.name)
      .addEditor(user.email);
  var body = doc.getBody();
  var table = [['Session', 'Date', 'Time', 'Location']];
  for (var i = 0; i < response.length; i++) {
    table.push([response[i][0], response[i][1].toLocaleDateString(),
        response[i][2].toLocaleTimeString(), response[i][4]]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: 'Thanks for registering! Here\'s your itinerary: ' + doc.getUrl(),
    attachments: doc.getAs(MimeType.PDF),
  });
}












function importEvents(e) {
  var calendar_name = e.parameter.calendar;
  var startDate = new Date(e.parameter.start);
  var endDate = new Date(e.parameter.end);
  var Calendar = CalendarApp.getCalendarsByName(calendar_name);
  var events = Calendar[0].getEvents(startDate, endDate);

  if (events[0]) {
    var eventarray = new Array();
    var line = new Array();
    line.push('Member Name','Mobile', 'Time');
    eventarray.push(line);
    for (var i = 0; i < events.length; i++) {
      line = new Array();
      FUS1=new Date(events[i]).toString().substr(25,6)+':00';
      line.push(events[i].getTitle());
      line.push(' ' + events[i].getDescription().replace(/[^\d]/g, ''));
      line.push(Utilities.formatDate(events[i].getStartTime(), FUS1, "MMM-dd-yyyy")+' at ' +Utilities.formatDate(events[i].getEndTime(), FUS1, "HH:mm"));
      eventarray.push(line);
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    sheet.getActiveRange().clear(); // get rid of whatever is already there.
    sheet.getRange(1,1,eventarray.length,eventarray[0].length).setValues(eventarray);
    sheet.setColumnWidth(1, 450);sheet.setColumnWidth(2, 150);sheet.setColumnWidth(3, 150);sheet.setColumnWidth(4, 250);sheet.setColumnWidth(5, 90);
    sheet.setFrozenRows(1);
  } else {
    var startstring = Utilities.formatDate(e.parameter.start, FUS1, "MMM-dd-yyyy");
    var endstring = Utilities.formatDate(e.parameter.end, FUS1, "MMM-dd-yyyy");
    Browser.msgBox('There are no events in the calendar between ' + startstring + ' and ' + endstring + ' in calendar '+calendar_name);
  }
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function Cal_to_sheet() {
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle('Calendar Import').setHeight('320').setWidth('440').setStyleAttribute('background','beige');
  // Create a grid with 3 text boxes and corresponding labels
  var grid = app.createGrid(3, 2);
  grid.setWidget(0, 0, app.createLabel('Calendar:').setWidth('100'));
  var list = app.createListBox();
  list.setName('calendar');
  grid.setWidget(0, 1, list);
  var calendars = CalendarApp.getAllCalendars();
  for (var i = 0; i < calendars .length; i++) {
    list.addItem(calendars[i].getName());
  }
  grid.setWidget(1, 0, app.createLabel('Start date:').setWidth('100'));
  grid.setWidget(1, 1, app.createDateBox().setId("start"));
  grid.setWidget(2, 0, app.createLabel('End date :').setWidth('100'));
  grid.setWidget(2, 1, app.createDateBox().setId("end"));
  var panel = app.createVerticalPanel();
  panel.add(grid);
  var button = app.createButton('Import');
  var handler = app.createServerHandler("importEvents");
  handler.addCallbackElement(grid);
  button.addClickHandler(handler);
  panel.add(button);
  app.add(panel);
  doc.show(app);
}
