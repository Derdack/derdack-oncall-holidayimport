/*
This script reads a Google calendar ICS file (e. g. https://calendar.google.com/calendar/ical/en.german%23holiday%40group.v.calendar.google.com/public/basic.ics), and transfers holidays from it into the OnCallPlanHolidays table of Enterprise Alert.

Please set up your STRING_DB_CONNECTION first. Ask your database administrator for support if you do not know it.

Then please set up your STRING_COUNTRY (e. g. "Berlin" or "Brandenburg" etc.) A warning appears if it may not work.

At least mark your team name(s) with the STRING_TEAMS constant. Leave it empty, if you want to save the holidays in all calendars.

Afterwards please execute this script by double-clicking or calling it directly from the command line by typing "WScript CalendarImporter.js".

By default the dry-run marker is activated. If you are happy with your prepared settings (DB, country, team names) please set BOOL_DRY_RUN to false to arm this script. Then re-execute it.

For further support contact us at support@de.derdack.com

v1.0.0 (06.09.2018, Frank Gutacker)

Copyright 2018 Derdack GmbH, www.derdack.com, Enterprise Alert is a registered trademark of Derdack GmbH
*/


var STRING_DB_CONNECTION 	= "Driver=SQL Server Native Client 11.0;Server=sqlserver.derdack-support.local;Trusted_Connection=No;UID=sa;PWD=Derdack!;Database=EnterpriseAlert2017"
var STRING_COUNTRY			= "Brandenburg"; // Name of country (e. g. "Baden-Württemberg")
var STRING_TEAMS			= "Support"; // empty string for all, or comma-separated list of team names (e. g. "Administrators,Standard User")

var STRING_ICS_FILE 		= ".\\basic.ics";

// Please set this value to false to store holidays in database
var BOOL_DRY_RUN			= false;

var oFs 	= new ActiveXObject("Scripting.FileSystemObject");
var oDb 	= new ActiveXObject("ADODB.Connection");

// Database helper
var DB = { 
	close: function() {
		try {
			oDb.Close();
		} catch(e) {
			// noop
		}	
	},
	
	// Executes SQL, and returns number of affected rows.
	count: function(sSql) {
		DB.open();
		var i = 0;
		try {
			var oRes = oDb.Execute(sSql);
			if (oRes.EOF) {
				return 0;
			}
			while (!oRes.EOF) {
				oRes.MoveNext();
				i++;
			}
		} catch(e) {
			e.message = "DB.count: " + (e.message ? e.message : e);
			throw e;
		}
		DB.close();
		
		return i;
	},
	
	// Return holiday ID from OnCallPlanHolidays table for given team ID, and date string
	getHolidayId: function(iTeamId, sDate) {
		DB.open();
		var sId = null;
		try {
			var oRes = oDb.Execute("SELECT ID FROM OnCallPlanHolidays WHERE OnCallPlanID=" + iTeamId + " AND Holiday='" + sDate + "'");
			if (!oRes.EOF) {
				sId = oRes.Fields.Item('ID').Value;
			}
		} catch(e) {
			e.message = "DB.getHolidayId: " + (e.message ? e.message : e);
			throw e;
		}
		DB.close();
		
		return sId;
	},
	
	// Return array of team IDs for given comma-separated list of team names
	getTeamIds: function(sTeams) {
		DB.open();
		var aTeams = sTeams.split(",");
		sTeams = "'" + aTeams.join("','") + "'";
		var i = 0;
		var aIds = [];
		try {
			var oRes = oDb.Execute("SELECT TeamID FROM TeamOnCallPlans" + (sTeams != "''" ? " WHERE TeamDisplayName IN (" + sTeams + ")" : ""));
			while (!oRes.EOF) {
				aIds[i++] = oRes.Fields.Item('TeamID').Value;
				oRes.MoveNext();
			}
		} catch(e) {
			e.message = "DB.getTeamIds: " + (e.message ? e.message : e);
			throw e;
		}
		DB.close();
		
		return aIds;
	},
	
	// Insert holiday in OnCallPlanHolidays table for given team ID, and date string
	insertHoliday: function(iTeamId, sDate) {
		DB.open();
		var oRes = false;
		try {
			oRes = oDb.Execute("INSERT INTO OnCallPlanHolidays(OnCallPlanID, Holiday) VALUES(" + iTeamId + ", '" + sDate + "')");
		} catch(e) {
			e.message = "DB.insertHoliday: " + (e.message ? e.message : e);
			throw e;
		}
		DB.close();
		
		return oRes;
	},

	open: function() {
		DB.close();
		try {
			oDb.Open(STRING_DB_CONNECTION);
		} catch(e) {
			WScript.echo("!!! Database not reachable (" + e.message ? e.message : e + ") !!!");
			throw(e);
		}
	}
}

function getNext(sLine, aContent, i) {
	for (var k = i; k < aContent.length; k++) {
		if (aContent[k].indexOf("END:VEVENT") != -1) {
			return "";
		}
		if (aContent[k].indexOf(sLine) != -1) {
			if (aContent[k].indexOf("\\") != -1) {
				var sResult = "";
				while (aContent[k].indexOf("\\") != -1) {
					sResult += aContent[k++];
				}
				return sResult.split(":")[2].replace(/\\/g, "");
			} else {
				if (aContent[k].indexOf("DESCRIPTION") != -1) {
					return aContent[k].split(":")[2];
				}
				return aContent[k].split(":")[1];
			}
		}
	}
}

// Save holiday. But check first, if it really is. If not so, or if it could not get saved, return the date string for further processing.
function saveHoliday(iTeamId, sDate, sName) {
	if (sName.indexOf("Advent") != -1) {
		return sDate;
	}
	if (sName.indexOf("Carnival") != -1) {
		return sDate;
	}
	if (sName.indexOf("Daylight Saving Time") != -1) {
		return sDate;
	} 
	if (sName.indexOf("Eve") != -1) {
		return sDate;
	}
	if (sName.indexOf("International Women's Day") != -1) {
		return sDate;
	}
	if (sName.indexOf("Maundy Thursday") != -1) {
		return sDate;
	}
	if (sName.indexOf("Shrove") != -1) {
		return sDate;
	}
	if (sName.indexOf("Valentine's Day") != -1) {
		return sDate;
	}
	var sNotSaved = null;
	if (!DB.getHolidayId(iTeamId, sDate)) {
		if (!DB.insertHoliday(iTeamId, sDate)) {
			sNotSaved = sDate;
		}
	}
	
	return sNotSaved; 
}

// Check if teams from STRING_TEAMS are existing
function hasTeams(STRING_TEAMS) {
	if (!STRING_TEAMS) {
		return true;
	}
	var aTeams = STRING_TEAMS.split(",");
	var sTeams = "'" + aTeams.join("','") + "'";
	var iRes = DB.count("SELECT * FROM TeamOnCallPlans WHERE TeamDisplayName IN (" + sTeams + ")");
	
	return (iRes == aTeams.length);
}

// Import holidays from array of date objects, containing date, name, and info values.
function importHolidays(aEntries) {
	var aTeamIds = DB.getTeamIds(STRING_TEAMS);
	var aNotSaved = [];
	for (var i = 0, m = 0; i < aTeamIds.length; i++) {
		for (var k = 0; k < aEntries.length; k++) {
			if (!STRING_COUNTRY || !aEntries[k].info || aEntries[k].info.indexOf(STRING_COUNTRY) != -1) {
				var sNotSaved = saveHoliday(aTeamIds[i], aEntries[k].date.substr(0,4) + "-" + aEntries[k].date.substr(4,2) + "-" + aEntries[k].date.substr(6,2), aEntries[k].name)
				if (sNotSaved) {
					aNotSaved[m++] = sNotSaved;
				}
			}
		}
	}
	if (aNotSaved.length > 0) {
		WScript.echo("Dates not saved: " + aNotSaved.join(", "));
	} else {
		WScript.echo("Done");
	}
}

// Delete all holidays of all teams
function deleteAllHolidays() {
	
	oDb.Open(STRING_DB_CONNECTION);

	oDb.Execute("DELETE OnCallPlanHolidays;");
	
	oDb.Close();
	
	WScript.echo("All holidays deleted.");
}

// Fetch, and check content from STRING_ICS_FILE. Save holidays in database.
function main() {
	var sContent = FS.file2string(STRING_ICS_FILE);
	
	var aCalendarEntries = [];
	
	var aContent = sContent.split("\n");
	for (var i = 0, k = 0; i < aContent.length; i++) {
		if (aContent[i].indexOf("BEGIN:VEVENT") != -1) {
			var sDate = getNext("DTSTART", aContent, i);
			var sSummary = getNext("SUMMARY", aContent, i);	
			var sDescription = getNext("DESCRIPTION", aContent, i);	
			aCalendarEntries[k++] = {
				"date" : sDate,
				"name" : sSummary.replace(/ \(.*\)/, ""),
				"info" : sDescription ? sDescription.replace(/[\s*\\]/g, "") : ""
			}
		}
	}
	
	var sAll = "";
	var aCountries = [];
	for (i = 0; i < aCalendarEntries.length; i++) {
		aTempCountries = aCalendarEntries[i].info.split(",");
		var bHasCountry = false;
		for (var k = 0; k < aTempCountries.length; k++) {
			bHasCountry = false;
			for (var m = 0; m < aCountries.length; m++) {
				if (aTempCountries[k] == aCountries[m]) {
					bHasCountry = true;
					break;
				}
			}
			if (!bHasCountry && aTempCountries[k]) {
				aCountries[aCountries.length] = aTempCountries[k];
			}
		}
	}
	
	if (STRING_COUNTRY && aCountries.join(", ").indexOf(STRING_COUNTRY) == -1) {
		WScript.echo("Country \"" + STRING_COUNTRY + "\" not found.\n\nPlease choose one of those:\n\n" + aCountries.join(", ") + "\n\nor an empty string (STRING_COUNTRY = \"\") if you do not want to import local holidays.");
	} else {
		if (!hasTeams(STRING_TEAMS)) {
			WScript.echo("Please check your teams. One ore more of them do not exist.");
		} else {
			WScript.echo(
				aCalendarEntries.length + " calendar entries found in ICS file. " + 
				"Importing public holidays" + (STRING_COUNTRY ? ", and local holidays of " + STRING_COUNTRY : "") + ". " + 
				"Affected teams: " + (STRING_TEAMS ? STRING_TEAMS : "all teams") + "." + 
				(BOOL_DRY_RUN ? " (Dry run only - no changes will be made)" : "")
			);
			if (!BOOL_DRY_RUN) {
				importHolidays(aCalendarEntries);
			}
		}
	}
}

var FS = {
	// Transfer file content into one string (lines separated by line break).
	file2string: function(sFile) {
		var sRes = "";

		try {
			var oFileHandler = oFs.OpenTextFile(sFile, 1);

			for (var i = 0; !oFileHandler.AtEndOfStream; i++) {
				sRes += oFileHandler.ReadLine() + "\n";
			}

			oFileHandler.Close();
		} catch (e) {
			if (oFileHandler) { 
				oFileHandler.Close(); 
			}
			WScript.echo("FS.file2string() error message: " + (e.message ? e.message : e));
		}
		
		return sRes;
	}
}

// Trigger main method
main();

// Helper to delete all holidays
//deleteAllHolidays();
