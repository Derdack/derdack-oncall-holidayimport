# Holiday Import from ICS Files

In Enterprise Alert you can define public holidays for each team. The provided SQL script allows you to import holiday information from .ics files. You can find these files in sources like Google Calendar (e.g. https://calendar.google.com/calendar/ical/en.german%23holiday%40group.v.calendar.google.com/public/basic.ics).

In order to use the script you need to adapt the following parameters:

STRING_DB_CONNECTION: Database connection string to the Enterprise Alert database.
STRING_COUNTRY: Filter for the country or region if specifier in the .ics file.
STRING_TEAMS: The name of the team for which you would like to import the holidays, or empty in order to import the holidays for all teams.

The credits for this script go to my colleague Frank Gutacker.
