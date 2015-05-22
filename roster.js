// Link spreadsheet menu when loaded.
function onOpen() {
  var menu = [
    {name: 'Start new term', functionName: 'PopulateSheetFromCalendar'},
    {name: 'Update spreadsheet from calendar', functionName: 'runImport'},
    {name: 'Populate signup form (should happen automatically)', functionName: 'UpdateFormList'}
  ]
  SpreadsheetApp.getActive().addMenu('Rostering', menu);
}

var dropdownTitle = 'Roster Dates (choose one)';
var dropdownHelp = "Selecting a date with the text 'Key' indicates you will hold the keys.";

/**
 * Creates/updates a Google Form that allows respondents to select which roster
 * sessions they can make it to.
 */
function UpdateFormList() {
  var listItems = []
  var events = getEvents();
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle().match(unassignedRegexp)) {
      listItems.push(eventTimeLabel(events[i]));
    }
  }
  // If multiple timeslots have the same value,
  // the selection comes through as an array.
  // Prevent this by calling uniq.
  storedFormList()
    .setChoiceValues(uniq(listItems))
    .setTitle(dropdownTitle)
    .setRequired(true)
    .setHelpText(dropdownHelp);

    // These are the available dates for the forecoming term. Please choose one.
}

function uniq(array) {
  if (array == null) return [];
  var result = [];
  var seen = [];
  for (var i = 0, length = array.length; i < length; i++) {
    var value = array[i];
    if (result.indexOf(value) < 0) {
      result.push(value);
    }
  }
  return result;
};

var keyRegex = /Key\)?$/i;
var commRegex = /comm\)?$/i;

function eventTimeLabel(event) {
  label = event.getStartTime().toDateString();
  if (! event.getTitle().match(unassignedRegexp)) {
    label = Utilities.formatDate(event.getStartTime(), FUS1, "EEEEE, d MMMMM") + "(TAKEN)"
  }
  else if (event.getTitle().match(keyRegex)) {
    label = Utilities.formatDate(event.getStartTime(), FUS1, "EEEEE, d MMMMM") + " - Key Person";
  }
  else if (event.getTitle().match(commRegex)) {
    label = Utilities.formatDate(event.getStartTime(), FUS1, "EEEEE, d MMMMM") + " - Committee";
  }
  else
  {
    label = Utilities.formatDate(event.getStartTime(), FUS1, "EEEEE, d MMMMM")
  }
  return label;
}

function logMail(msg, details) {
  MailApp.sendEmail(
    "oakleigh.toylibrary@gmail.com",
    msg,
    JSON.stringify(details)
  );
}

/*
 *
 * Store/retrieve which spreadsheet we're using
 *
 */
function setSelectedSpreadSheet(sheet) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('SS_DOC_ID', sheet.getParent().getId());
  scriptProperties.setProperty('SS_DOC_SHEET_NAME', sheet.getName());
}

function storedSpreadSheet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  ssId = scriptProperties.getProperty('SS_DOC_ID');
  name = scriptProperties.getProperty('SS_DOC_SHEET_NAME');
  return SpreadsheetApp.openById(ssId).getSheetByName(name);
}

/*
 *
 * Store/retrieve which form we created
 *
 */
function storedForm() {
  var scriptProperties = PropertiesService.getScriptProperties();
  formId = scriptProperties.getProperty('ROSTERING_FORM_ID');
  var form
  try {
    form = FormApp.openById(formId);
  } catch (e) {
    // IDGAF
  }

  if (form) {
    return form
  }

  // Create the form and add a multiple-choice question for each timeslot.
  form = FormApp.create('Rostered Dates Form');

  ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create();

  form.addTextItem().setTitle('Full Name').setRequired(true);
  form.addTextItem().setTitle('Mobile number').setRequired(true);

  scriptProperties.setProperty('ROSTERING_FORM_ID', form.getId());
  return form;
}

function storedFormList() {
  var scriptProperties = PropertiesService.getScriptProperties();
  listId = scriptProperties.getProperty('GF_ROSTER_FORM_LISTID');

  var listItem
  try {
    listItem = storedForm().getItemById(listId).asListItem();
  } catch (e) {
    // IDGAF
  }
  if (listItem) {
    return listItem;
  }
  listItem = storedForm().addListItem()
    .setTitle(dropdownTitle)
    .setRequired(true)
    .setHelpText(dropdownHelp)
    .setChoiceValues(["None Yet"]);

  scriptProperties.setProperty('GF_ROSTER_FORM_LISTID', listItem.getId());

  return listItem;
}

/*
 *
 * Store/retrieve date range
 *
 */
function setStoredDateRange(start, end) {
  var scriptProperties = PropertiesService.getScriptProperties();
  // Force date-to-string conversion to use ints
  scriptProperties.setProperty('START_DATE_RANGE', "" + (1 * start));
  scriptProperties.setProperty('END_DATE_RANGE', "" + (1 * end));
}

function storedStartDate() {
  var scriptProperties = PropertiesService.getScriptProperties();
  dateStr = scriptProperties.getProperty('START_DATE_RANGE');
  // Convert string to int by multiplication because javascript.
  return new Date(1 * dateStr)
}

function storedEndDate() {
  var scriptProperties = PropertiesService.getScriptProperties();
  dateStr = scriptProperties.getProperty('END_DATE_RANGE');
  // Convert string to int by multiplication because javascript.
  return new Date(1 * dateStr)
}

/*
 *
 * Store/retrieve calendar deets
 *
 */
function setSelectedCalendar(calId) {
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('CALENDAR_ID', calId);
}

function rosterCalendar() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var calId = scriptProperties.getProperty('CALENDAR_ID');
  var result = CalendarApp.getCalendarById(calId);
  if (!result) {
    return null;
  }
  return result;
}

function getEvents() {
  return rosterCalendar().getEvents(storedStartDate(), storedEndDate());
}

var unassignedRegexp = /^Nobody/i;

/**
 * A trigger-driven function that updates the calendar after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  var responses = e.response.getItemResponses()
  var submission = {}
  for (var i = 0; i < responses.length; i++) {
    submission[responses[i].getItem().getTitle()] = responses[i].getResponse()
  }

  var user = {
    name: submission["Full Name"],
    mobile: submission["Mobile number"],
    timeslot: submission[dropdownTitle]
  };

  // Get a public lock on this script, because we're about to modify a shared resource.
  var lock = LockService.getPublicLock();
  // Wait for up to 5 seconds for other processes to finish.
  lock.waitLock(5000);
  try {
    var events = getEvents();
    var event = null;
    for (var i = 0; i < events.length; i++) {
      if (user.timeslot === eventTimeLabel(events[i])) {
        event = events[i]
        break; // Found it!
      }
    }
    if (!event) {
      logMail("Form submission couldn't find event!", [e, user]);
    } else {
      // Update the event
      if (event.getTitle().match(unassignedRegexp)) {
        if (event.getTitle().match(keyRegex)) {
          event.setTitle(user.name + " - Key");
        } else {
          event.setTitle(user.name);
        }
        // Valid AU mobile numbers only have digits in them.
        user.mobile = user.mobile.replace(/[^\d]/g, '');
        // Twilio wants +614XX, not 04XX for SMS.
        user.mobile = user.mobile.replace(/^04/, '+614');
        event.setDescription(user.mobile);
      } else {
        if (event.getTitle() === user.name) {
          // All is good
        } else {
          logMail("Whoops: Two people tried to book the same slot at once", [e, user, e.getTitle()]);
        }
      }
    }
  } finally {
    // Release the lock so that other processes can continue.
    lock.releaseLock();
    UpdateFormList();
    runImport();
  }
}

var FUS1=new Date().toString().substr(25,6)+':00';
function importEvents(e) {
  var startDate = new Date(e.parameter.start);
  var endDate = new Date(e.parameter.end);
  setStoredDateRange(startDate, endDate); // Record these to use elsewhere.

  var calendar_name = e.parameter.calendar;
  var Calendar = CalendarApp.getCalendarsByName(calendar_name);
  setSelectedCalendar(Calendar[0].getId()); // Record these to use elsewhere.

  var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  setSelectedSpreadSheet(currentSheet);

  runImport();
}

function runImport() {
  var startDate = storedStartDate();
  var endDate = storedEndDate();

  var events = getEvents();

  if (events[0]) {
    var eventarray = new Array();
    var line = new Array();
    line.push('Member Name','Mobile', 'Time');
    eventarray.push(line);
    for (var i = 0; i < events.length; i++) {
      line = new Array();
      FUS1=new Date(events[i]).toString().substr(25,6)+':00';
      line.push(events[i].getTitle());

      mobile = events[i].getDescription();
      line.push(mobile);
      line.push(Utilities.formatDate(events[i].getStartTime(), FUS1, "MMM-dd-yyyy")+' at ' +Utilities.formatDate(events[i].getEndTime(), FUS1, "HH:mm"));
      eventarray.push(line);
    }

    var sheet = storedSpreadSheet();
    try {
      // Attempt to focus, if we're not running interactively this will fail.
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(sheet.getName())
        .activate();
    } catch (e) {
    }

    sheet.getActiveRange().clear(); // get rid of whatever is already there.
    sheet.getRange(1,1,eventarray.length,eventarray[0].length).setValues(eventarray);
    sheet.setColumnWidth(1, 450);sheet.setColumnWidth(2, 150);sheet.setColumnWidth(3, 150);sheet.setColumnWidth(4, 250);sheet.setColumnWidth(5, 90);
    sheet.setFrozenRows(1);
    UpdateFormList();
  } else {
    var startstring = Utilities.formatDate(startDate, FUS1, "MMM-dd-yyyy");
    var endstring = Utilities.formatDate(endDate, FUS1, "MMM-dd-yyyy");
    Browser.msgBox('There are no events in the calendar between ' + startstring + ' and ' + endstring + ' in calendar ' + rosterCalendar().getName());
  }
}

function PopulateSheetFromCalendar() {
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

function InitialConfig() {
  PopulateSheetFromCalendar();
}
