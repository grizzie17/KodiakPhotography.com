<%
DIM	g_sSiteName
g_sSiteName = "Kodiak Photography"


DIM	g_sSiteTitle
g_sSiteTitle = "Kodiak Photography (a division of Bear Consulting Group)"


DIM g_sCookiePrefix
g_sCookiePrefix = "KDKP"


DIM	g_sDomain
g_sDomain = "KodiakPhotography.com"



DIM g_nPictureDelay
g_nPictureDelay = 3



DIM g_bGalleryRandomize
g_bGalleryRandomize = TRUE


' The BaseDataFolder defines where to find the "Gallery" and "Marquee" folders
' If this is empty then we search for the "cgi-bin" folder
' If specified; make sure that you provide a complete specification including drive.

DIM g_sBaseDataFolder
g_sBaseDataFolder = ""



DIM	g_sMailServer
g_sMailServer = "mail.KodiakPhotography.com"


DIM g_sAnnonUser
DIM g_sAnnonPW

g_sAnnonUser = "Notifications@" & g_sDomain
g_sAnnonPW = "bear1701"




DIM g_sCalendarHiddenList
g_sCalendarHiddenList = "" _
	&	"Email" & vbTAB & "email,email2" & vbTAB & "Email Events" & vbLF _
	&	"Newsletter" & vbTAB & "key,newsletter,c-activity,rally,district,safety,entertainment,funday,charity" & vbTAB & "Newsletter Events"








DIM g_nTimeZoneOffset


FUNCTION config_dateFromWeekNumber( nWeek, nWDay, nMonth, nYear )
	DIM	dFirst
	DIM	nFirstDay
	DIM	x
	dFirst = DATEVALUE(nMonth & "/1/" & nYear )
	nFirstDay = WEEKDAY(dFirst)
	x = 1 + ( nWDay - nFirstDay + 7 ) MOD 7
	x = x + (7 * (nWeek - 1))
	config_dateFromWeekNumber = DATEVALUE( nMonth & "/" & x & "/" & nYear )
	'Response.Write "dateFromWeekNumber = " & config_dateFromWeekNumber & "<br>"
END FUNCTION

SUB config_adjustDaylightSavingsTime()
	DIM	nMonth
	DIM	nYear
	nMonth = MONTH(NOW)
	nYear = YEAR(NOW)
	IF 3 < nMonth  AND  nMonth < 11 THEN
		g_nTimeZoneOffset = g_nTimeZoneOffset + 1	' Daylight Savings Time
	ELSEIF 3 = nMonth THEN
		'second sunday
		IF 0 < DATEDIFF("h",config_dateFromWeekNumber(2, 1, 3, nYear), NOW) THEN
			g_nTimeZoneOffset = g_nTimeZoneOffset + 1	' Daylight Savings Time
		END IF
	ELSEIF 11 = nMonth THEN
		'first sunday
		IF 0 > DATEDIFF("h",config_dateFromWeekNumber(1, 1, 11, nYear), NOW) THEN
			g_nTimeZoneOffset = g_nTimeZoneOffset + 1	' Daylight Savings Time
		END IF
	END IF
	
END SUB


IF LCASE(Request.ServerVariables("SERVER_NAME")) = "localhost" THEN
	g_nTimeZoneOffset = 0
ELSE
	g_nTimeZoneOffset = -6
	config_adjustDaylightSavingsTime
END IF



%>