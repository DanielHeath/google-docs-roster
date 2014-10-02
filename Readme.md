= Roster

Google docs based rostering, with Twilio for SMS via Zapier.

= Setup

Register a Google, Twilio and Zapier account.
Create a google docs spreadsheet, then add a script via Tools -> Script Editor -> Blank Script.
Paste in roster.js from this project and save, then re-load the spreadsheet.

A 'Roster' menu will appear at the top of the page.

In google calendar, create all the roster times as events.
If two people are needed at the same time, create two events.
The event title must start with 'Nobody'.

Click Roster -> Setup to
 * Import the calendar into the spreadsheet
 * Create a google form which you can send to your members

Filling in the form will alter the calendar and the spreadsheet.

Changes to the spreadsheet *WILL BE LOST* whenever someone fills in the form.

FIXME: Alter the script so that when the sheet changes it wipes things out.

Test the form created by this (go to google drive to find the form).

Once you fill out one entry, it should update the calendar *and* the spreadsheet.

Go to Zapier and create a "Google Calendar <-> Twilio SMS" integration.














# Copy in script -> setup linkage ?

    Choose Resources > Current project's triggers. You see a panel with the message No triggers set up. Click here to add one now.
    Click the link.
    Under Run, select the function you want executed by the trigger. (That's onFormSubmit(), in this case.)
    Under Events, select From Spreadsheet.
    From the next drop-down list, select On form submit.
    Click Save.
