
-- Get all on-call plan IDs (including auto-rotations) --
DECLARE @PeriodStartDate varchar(20) = '2018-01-01 00:00:00'; /* The start date / time of the search period */
DECLARE @PeriodEndDate varchar(20) = '2019-01-01 00:00:00'; /* The end date / time of the search period */
SELECT DISTINCT
TeamID, /* ID of the team */
TeamDisplayName, /* Name of the team */
HandOverTime, /* Hand-over time in seconds of the day */
ProfileID, /* Profile ID of the user */
ADRNAME, /* Name of the user */
PROFNAME, /* Username of the user */
r.ShiftStart, /* Start date of the duty */
r.ShiftEnd, /* End date / time of the duty */
cast( dateadd(second,Handovertime, '2000-01-01') as time) HandOverTime_T,
r.ShiftOptions,
case 
  when r.Hierarchy = 1 then 'Backup' /* Backup (level two) */
  when r.Hierarchy = 0 and ((r.ShiftOptions & 2) = 2) then 'Stand-In' /* Stand-In */
  when r.Hierarchy >= 2 then 'Escalation' /* Backup of a higher level than two */
  else 'Primary' /* Primary */
  end as Hierarchy_Type,
/* r.Hierarchy, */
ADRMOBILE, /* Mobile phone number of the user */
ADREMAIL, /* Email address of the user */
ADRSIP, /* SIP address of the user */
ADRPHONE, /* Phone number of the user */
ADRPAGER /* Pager address of the user */
/* , OnCallPlans.ID, OnCallPlanUsers.ID */
FROM OnCallPlans
INNER JOIN OnCallPlanUsers ON OnCallPlans.ID  = OnCallPlanUsers.OnCallPlanID
FULL OUTER JOIN OnCallPlanShifts ON (OnCallPlanUsers.ID = OnCallPlanShifts.OnCallPlanUserID)
INNER JOIN TeamOnCallPlans ON (OnCallPlans.ID = TeamOnCallPlans.OnCallPlanID)
INNER JOIN MMPROFILES ON (OnCallPlanUsers.ProfileID = MMPROFILES.ID)    
CROSS APPLY dbo.fn_OnCallPlanUserGetDuties(OnCallPlans.ID, OnCallPlanUsers.ID, @PeriodStartDate, @PeriodEndDate, (OnCallPlans.Options & 1), OnCallPlans.HandOverTime, 7) as r
WHERE
	r.Hierarchy <= 1 /* Only display level-two backups */
	/* AND TeamDisplayName = 'Support' */ /* Filter for a specific team */
ORDER BY TeamDisplayName ASC;
