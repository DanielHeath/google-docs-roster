= Roster

Google docs based rostering, with Twilio for SMS via Zapier.

= Setup

Register a Google, Twilio and Zapier account.
Create a google docs spreadsheet, then add a script via Tools -> Script Editor -> Blank Script.
Paste in roster.js from this project and save.

In google calendar, create all the roster times as events (in your default calendar).
If two people are needed on duty at the same time, create two events.
The event title must be 'Nobody' or 'Nobody - Key'.

Re-load the spreadsheet; a 'Roster' menu will appear at the top of the page.
Click Roster -> 'Start new term' in this menu, then select
the start & end dates for the roster term.
This will fill the spreadsheet with the entries in the calendar, and
create a form in your google drive account.

Filling in the form will alter the calendar and the spreadsheet.

Changes to the spreadsheet *WILL BE LOST* whenever someone fills in the form.

Test the form created by this (go to google drive to find the form).

Once you fill out one entry, it should update the calendar *and* the spreadsheet.

Go to Zapier and create a "Google Calendar <-> Twilio SMS" integration to send SMS when an event is coming up. TODO: screenshots of how that works.














# Copy in script -> setup linkage ?

    Choose Resources > Current project's triggers. You see a panel with the message No triggers set up. Click here to add one now.
    Click the link.
    Under Run, select the function you want executed by the trigger. (That's onFormSubmit(), in this case.)
    Under Events, select From Spreadsheet.
    From the next drop-down list, select On form submit.
    Click Save.
