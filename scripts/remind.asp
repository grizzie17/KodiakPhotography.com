<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2001 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------
%>
<!--#include file="dateutils.asp"-->
<!--#include file="htmlformat.asp"-->
<%

CONST kTropicalYear = 365.242190


DIM	g_remind_handleSubsText
g_remind_handleSubsText = "remind_handleSubsText"

g_htmlFormat_metaSubsFunc = "remind_metaSubsFunc"
g_htmlFormat_pictureFunc = "remind_pictureFile"

	
	
SUB appendList( n, a, v )
	IF UBOUND(a) < n+1 THEN REDIM PRESERVE a(n+10)
	a(n) = v
	n = n + 1
END SUB

FUNCTION findFolder( sFolderName )
	DIM	sTemp
	DIM	oFolder

	findFolder = ""
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))
	DO UNTIL oFolder.IsRootFolder
		sTemp = g_oFSO.BuildPath( oFolder.Path, sFolderName )
		IF g_oFSO.FolderExists( sTemp ) THEN
			findFolder = sTemp
			EXIT DO
		END IF
		ON Error RESUME Next
		SET oFolder = oFolder.ParentFolder
		IF Err THEN EXIT DO
	LOOP
	ON Error GOTO 0
	SET oFolder = Nothing

END FUNCTION


DIM	g_sPagePicturePath
DIM	g_sPagePictureVirtPath
g_sPagePicturePath = ""
g_sPagePictureVirtPath = ""

' this function depends on the global "sName" to define the faculty name-id
FUNCTION remind_pictureFile( sLabel )
	DIM	sFile
	IF 0 = LEN(g_sPagePicturePath) THEN
		DIM	sProfilePath
		DIM	sPicturePath

		sPicturePath = findFolder( "images\remind" )
		IF 0 < LEN(sPicturePath) THEN
			DIM	sRootPath
			g_sPagePicturePath = sPicturePath
			sRootPath = Server.MapPath("/")
			sPicturePath = MID(g_sPagePicturePath,LEN(sRootPath)+1)
			g_sPagePictureVirtPath = REPLACE(sPicturePath,"\","/",1,-1,vbTextCompare) & "/"
		END IF
	END IF
	sFile = sLabel & ".gif"
	IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN
		sFile = sLabel & ".jpg"
		IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN
			sFile = sLabel & ".png"
			IF NOT g_oFSO.FileExists( g_oFSO.BuildPath( g_sPagePicturePath, sFile )) THEN sFile = ""
		END IF
	END IF
	IF 0 < LEN(sFile) THEN
			sFile = g_sPagePictureVirtPath & sFile
	END IF
	remind_pictureFile = sFile
END FUNCTION




FUNCTION f_year( aArgs(), nDate, nTime )
		IF 0 < UBOUND(aArgs) THEN
			IF ISNUMERIC( aArgs(1) ) THEN
				DIM	yy
				DIM	d, m, y
				DIM	s, t
				
				yy = CINT(aArgs(1))
				gregorianFromJDate d, m, y, nDate
				d = y - yy
				s = CSTR(d)
				t = RIGHT(s,2)
				IF 1 < LEN(s)  AND  "1" = LEFT(t,1) THEN
					s = s & "th"
				ELSE
					SELECT CASE RIGHT(s,1)
					CASE "1"
						s = s & "st"
					CASE "2"
						s = s & "nd"
					CASE "3"
						s = s & "rd"
					CASE ELSE
						s = s & "th"
					END SELECT
				END IF
				f_year = s
			ELSE
				f_year = "{{invalid year specified}}"
			END IF
		ELSE
			f_year = "{{year not specified}}"
		END IF
END FUNCTION
	
FUNCTION f_days( aArgs(), nDate, nTime )
		DIM	n
		
		n = nDate - jdateFromVBDate( NOW )
		f_days = CSTR(n)
END FUNCTION
	
FUNCTION f_time( aArgs(), nDate, nTime )
		DIM	hh, mm, ss
		DIM	n
		DIM	sTemp

		n = FIX(nTime * CDBL(3600*24))
		hh = INT( n / 3600 )
		n = n - (hh*3600)
		mm = INT( n / 60 )

		sTemp = ""
		IF 0 < hh THEN
			IF hh < 10 THEN
				sTemp = "0" & hh & ":"
			ELSE
				sTemp = "" & hh & ":"
			END IF
		ELSE
			sTemp = "00:"
		END IF
		IF 0 < mm THEN
			IF mm < 10 THEN
				sTemp = sTemp & "0" & mm
			ELSE
				sTemp = sTemp & mm
			END IF
		ELSE
			sTemp = sTemp & "00"
		END IF
		
		f_time = sTemp
END FUNCTION
	
FUNCTION f_weekdays( aArgs(), nDate, nTime )
		DIM	n
		DIM	bInclusive
		DIM	jdNOW
		IF 0 < UBOUND(aArgs) THEN
			SELECT CASE LCASE(aArgs(1))
			CASE "include", "inclusive", "inclussive"
				bInclusive = TRUE
			CASE ELSE
				bInclusive = FALSE
			END SELECT
		ELSE
			bInclusive = FALSE
		END IF
		jdNOW = jdateFromVBDate( NOW() )
		n = diffWeekdaysJDates( jdNOW, nDate, bInclusive )
		f_weekdays = CSTR(n)
END FUNCTION
	

DIM	g_remind_nDate
DIM	g_remind_nTime

FUNCTION remind_metaSubsFunc( aArgs, sText )
	DIM	nDate
	DIM	nTime
	
	IF "" = sText THEN
		nDate = g_remind_nDate
		nTime = g_remind_nTime
		SELECT CASE LCASE( aArgs(0) )
		CASE "year", "years", "anniv"
			remind_metaSubsFunc = f_year( aArgs, nDate, nTime )
		CASE "day", "days"
			remind_metaSubsFunc = f_days( aArgs, nDate, nTime )
		CASE "weekdays", "weekday", "workdays", "workdays"
			remind_metaSubsFunc = f_weekdays( aArgs, nDate, nTime )
		CASE "time"
			remind_metaSubsFunc = f_time( aArgs, nDate, nTime )
		CASE ELSE
			remind_metaSubsFunc = ""
		END SELECT
	ELSE
		remind_metaSubsFunc = ""
	END IF
END FUNCTION

	
FUNCTION remind_handleSubsText( sText, nDate, nTime )
	DIM	sTemp
	g_remind_nDate = nDate
	g_remind_nTime = nTime
	sTemp = htmlFormatEncodeText( sText )
	IF 0 < INSTR(1,sText,"{{",vbTextCompare) THEN sTemp = htmlFormatMeta( sTemp )
	remind_handleSubsText = sTemp
END FUNCTION






'==== CCalendar ========================= CCalendar =====
PUBLIC aSubsSplit

CLASS CCalendar

	DIM	m_oXML
	DIM	m_oRoot
	
	DIM	m_nDate
	DIM	m_nTime
	DIM	m_sSubject
	DIM	m_sBody
	DIM	m_sStyle
	DIM	m_sCategories
	DIM	m_sID

	SUB Class_Initialize()
		m_nDate = 0
		m_sSubject = ""
		m_sBody = ""
		m_sCategories = ""
		m_sID = ""
		
		SET m_oXML = Server.CreateObject("msxml2.DOMDocument")
		'SET m_oXML = Server.CreateObject("Microsoft.XMLDOM")
		SET m_oRoot = m_oXML.createNode( 1, "calendar", "" )
		m_oXML.appendChild( m_oRoot )
	END SUB
	
	SUB Class_Terminate()
		SET m_oRoot = Nothing
		SET m_oXML = Nothing
	END SUB
	
	
	PROPERTY LET juliandate( nDate )
		m_nDate = FIX(nDate)
		m_nTime = nDate - m_nDate
	END PROPERTY
	
	PROPERTY LET subject( sSubject )
		m_sSubject = sSubject
	END PROPERTY
	
	PROPERTY LET body( sBody )
		m_sBody = sBody
	END PROPERTY
	
	PROPERTY LET style( sStyle )
		m_sStyle = sStyle
	END PROPERTY
	
	PROPERTY LET category( sCategory )
		IF "" = m_sCategories THEN
			m_sCategories = sCategory
		ELSE
			m_sCategories = m_sCategories & vbTAB & sCategory
		END IF
	END PROPERTY
	
	PROPERTY LET id( sID )
		m_sID = sID
	END PROPERTY
	
	PROPERTY GET xml
		xml = m_oXML.xml
	END PROPERTY
	
	PROPERTY GET xmldom
		SET xmldom = m_oXML
	END PROPERTY
	
	SUB outputMessage
		DIM	oEvent
		DIM	oDate
		DIM	oStyle
		DIM	oSubject
		DIM	oBody
		DIM	oText
		DIM	oAttr
		
		SET oEvent = m_oXML.createNode( 1, "event", "" )
		SET oText = m_oXML.createTextNode( logDateFromJDate( m_nDate ) )
		SET oDate = m_oXML.createNode( 1, "date", "" )
		
		DIM	dd, mm, yy
		gregorianFromJDate dd, mm, yy, m_nDate
		oDate.setAttribute "d", dd
		oDate.setAttribute "m", mm
		'oDate.setAttribute "mname", monthname(mm,3)
		oDate.setAttribute "y", yy
		oDate.setAttribute "wd", weekdayFromJDate( m_nDate )

		oDate.appendChild( oText )
		oEvent.appendChild( oDate )
		
		IF 0 < LEN( m_sStyle ) THEN
			SET oText = m_oXML.createTextNode( m_sStyle )
			SET oStyle = m_oXML.createNode( 1, "style", "" )
			oStyle.appendChild( oText )
			oEvent.appendChild( oStyle )
		END IF
		SET oText = m_oXML.createCDATASection( handleSubsText(m_sSubject, m_nDate, m_nTime) )
		SET oSubject = m_oXML.createNode( 1, "subject", "" )
		oSubject.appendChild( oText )
		oEvent.appendChild( oSubject )
		
		IF 0 < LEN( m_sBody ) THEN
			SET oText = m_oXML.createCDATASection( handleSubsText(m_sBody, m_nDate, m_nTime) )
			SET oBody = m_oXML.createNode( 1, "bodytext", "" )
			oBody.appendChild( oText )
			oEvent.appendChild( oBody )
		END IF
		
		IF 0 < LEN( m_sCategories ) THEN
			DIM	aSplitCat
			DIM	oCat
			DIM	sCat
			aSplitCat = SPLIT(m_sCategories,vbTAB)
			FOR EACH sCat IN aSplitCat
				SET oText = m_oXML.createTextNode( sCat )
				SET oCat = m_oXML.createNode( 1, "category", "" )
				oCat.appendChild( oText )
				oEvent.appendChild( oCat )
			NEXT 'sCat
			SET oCat = Nothing
			SET oText = Nothing
		END IF
		m_oRoot.appendChild( oEvent )
		
		SET oEvent = Nothing
		SET oDate = Nothing
		SET oSubject = Nothing
		SET oBody = Nothing
		SET oText = Nothing
		m_nDate = 0
		m_nTime = 0
		m_sSubject = ""
		m_sBody = ""
		m_sStyle = ""
		m_sCategories = ""
		m_sID = ""
	END SUB
	
	
	PRIVATE FUNCTION handleSubsText( sText, nDate, nTime )
		IF "" <> g_remind_handleSubsText THEN
			handleSubsText = EVAL( g_remind_handleSubsText & "( sText, nDate, nTime )" )
		ELSE
			handleSubsText = sText
		END IF
	END FUNCTION

END CLASS 'CCalendar





	PUBLIC aSplitDate
	PUBLIC aSplitTemp

'==== CCalendarFile ========================= CCalendarFile =====
CLASS CCalendarFile

	PRIVATE m_oXML
	PRIVATE m_sFilename
	
	PRIVATE	m_sCategories
	PRIVATE	m_sCategoryFilter
	PRIVATE	m_nDateBegin	'julian
	PRIVATE	m_nDateEnd		'julian
	PRIVATE	m_nPending		'number of days before now
	
	PRIVATE	m_nYearBegin
	PRIVATE	m_nYearNow
	
	SUB Class_Initialize()
		'SET m_oXML = Server.CreateObject("msxml2.DOMDocument")
		SET m_oXML = Server.CreateObject("Microsoft.XMLDOM")
		m_sFilename = ""
		
		m_sCategories = ""	'all
		m_sCategoryFilter = ""
		m_nDateBegin = 0	'all
		m_nDateEnd = 999999999
		m_nPending = 0
		
		m_nYearBegin = 0
		m_nYearNow = Year( Now )
	END SUB
	
	SUB Class_Terminate()
		SET m_oXML = Nothing
	END SUB
	
	
	PROPERTY LET file( sFilename )
		m_sFilename = sFilename
	END PROPERTY
	
	PROPERTY LET datebegin( nDateBegin )
		DIM	dd, mm, yy
		IF 0 < nDateBegin THEN
			m_nDateBegin = nDateBegin
			gregorianFromJDate dd, mm, yy, nDateBegin
			m_nYearBegin = yy
		ELSE
			m_nDateBegin = 0
			m_nYearBegin = 0
		END IF
	END PROPERTY
	
	PROPERTY LET dateend( nDateEnd )
		DIM	n
		n = CLNG(nDateEnd)
		IF 0 < n THEN
			m_nDateEnd = n
		ELSE
			pending = ABS( n )
		END IF
	END PROPERTY
	
	PROPERTY LET pending( nDays )
		m_nPending = ABS(CLNG(nDays))
		'm_nDateEnd = jdateFromVBDate( NOW )
		'm_nDateBegin = m_nDateEnd - m_nPending
		m_nDateBegin = jdateFromVBDate( NOW )
		m_nDateEnd = m_nDateBegin + m_nPending - 1
	END PROPERTY
	
	PRIVATE SUB buildCategoriesFilter( sCategories, oper )
		IF 0 < LEN(sCategories) THEN
			m_sCategories = sCategories
			aSplitDate = SPLIT( sCategories, ",", -1, vbTextCompare )
			DIM	sTemp, sBuild
			IF UBOUND(aSplitDate) < 1 THEN
				IF sCategories = "none" THEN
					m_sCategoryFilter = "[ not(category) ]"
				ELSE
					m_sCategoryFilter = "[ $any$ category = """ & sCategories & """]"
				END IF
			ELSE
				sBuild = ""
				FOR EACH sTemp IN aSplitDate
					IF sTemp = "none" THEN
						IF 0 < LEN(sBuild) THEN
							sBuild = sBuild & " " & oper & " not(category)"
						ELSE
							sBuild = "not(category)"
						END IF
					ELSE
						IF 0 < LEN(sBuild) THEN
							sBuild = sBuild & " " & oper & " ( $any$ category = """ & sTemp & """)"
						ELSE
							sBuild = "($any$ category = """ & sTemp & """)"
						END IF
					END IF
				NEXT 'sTemp
				m_sCategoryFilter = "[ " & sBuild & " ]"
			END IF
			'Response.Write "category-filter: " & m_sCategoryFilter
		ELSE
			m_sCategories = ""
			m_sCategoryFilter = ""
		END IF
	END SUB
	
	PROPERTY LET categories( sCategories )
		buildCategoriesFilter sCategories, "or"
	END PROPERTY
	PROPERTY LET categoriesUnion( sCategories )
		buildCategoriesFilter sCategories, "or"
	END PROPERTY
	PROPERTY LET categoriesIntersect( sCategories )
		buildCategoriesFilter sCategories, "and"
	END PROPERTY
	
	
	PRIVATE SUB dateSingle_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		nJulian = jdateFromGregorian( aSplitDate(2), aSplitDate(1), aSplitDate(0) )
		IF nEarly <= nJulian  AND  nJulian <= nLate THEN
			appendList nCount, aDates, nJulian
		END IF
	END SUB
	
	PRIVATE SUB dateMonthlyDayN_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	dd, mm, cm, cc, yy
		DIM	d, m, y

		gregorianFromJDate d, mm, yy, nEarly
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		cc = CINT(aSplitDate(0))
		cm = CINT(aSplitDate(1))
		dd = CINT(aSplitDate(2))
		IF 1 < cc THEN
			m = mm MOD cc
			m = mm - m + cm
			mm = m
		END IF
		DO
			DO WHILE mm < 13
				nJulian = jdateFromGregorian( dd, mm, yy )
				'Response.Write "julian = " & nJulian & "<br>" & vbCRLF
				IF nLate < nJulian THEN EXIT SUB
				IF 2 = mm  AND  28 < dd THEN
					gregorianFromJDate d, m, y, nJulian
					IF m <> mm  OR  d <> dd THEN nJulian = 0
				END IF
				IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
				mm = mm + cc
			LOOP
			mm = mm MOD 12
			yy = yy + 1
		LOOP
	END SUB
	
	PRIVATE SUB dateMonthlyWDay_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy, cc, cm, mm, wn, wd
		DIM	d, m, y

		gregorianFromJDate d, mm, yy, nEarly
		'Response.Write "date= " & oDate.text & ", year-begin= " & m_nYearBegin & ", yy= " & yy & ", nEarly= " & nEarly & ", nLate= " & nLate & "<br>" & vbCRLF
		'IF 0 = m_nYearBegin THEN
		'	yy = m_nYearNow
		'ELSE
		'	yy = m_nYearBegin
		'END IF
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		cc = CINT(aSplitDate(0))
		cm = CINT(aSplitDate(1))
		wn = CINT(aSplitDate(2))
		wd = CINT(aSplitDate(3))
		IF 1 < cc THEN
			m = mm MOD cc
			m = mm - m + cm
			mm = m
		END IF
		DO
			DO WHILE mm < 13
				nJulian = jdateFromWeeklyGregorian( wn, wd, mm, yy )
				'Response.Write "wn=" & wn & ", wd=" & wd & ", mm=" & mm & ", julian = " & nJulian & "<br>" & vbCRLF
				IF nLate < nJulian THEN EXIT SUB
				IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
				mm = mm + cc
			LOOP
			mm = mm MOD 12
			yy = yy + 1
		LOOP
	END SUB

	PRIVATE SUB dateMonthly_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("monthly")
		CASE "dayn"
			dateMonthlyDayN_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "wday"
			dateMonthlyWDay_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB

	PRIVATE SUB dateYearlyDayN_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	dd, mm, yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		'IF 0 < UBOUND(aSplitDate) THEN
		mm = CINT(aSplitDate(0))
		dd = CINT(aSplitDate(1))
		DO
			nJulian = jdateFromGregorian( dd, mm, yy )
			'Response.Write "julian = " & nJulian & "<br>" & vbCRLF
			IF nLate < nJulian THEN EXIT DO
			IF 2 = mm  AND  28 < dd THEN
				gregorianFromJDate d, m, y, nJulian
				IF m <> mm  OR  d <> dd THEN nJulian = 0
			END IF
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
		'END IF
	END SUB
	
	PRIVATE SUB dateYearlyWDay_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy, mm, wn, wd
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		'Response.Write "date= " & oDate.text & ", year-begin= " & m_nYearBegin & ", yy= " & yy & ", nEarly= " & nEarly & ", nLate= " & nLate & "<br>" & vbCRLF
		'IF 0 = m_nYearBegin THEN
		'	yy = m_nYearNow
		'ELSE
		'	yy = m_nYearBegin
		'END IF
		aSplitDate = SPLIT( oDate.text, " ", -1, vbTextCompare )
		mm = CINT(aSplitDate(0))
		wn = CINT(aSplitDate(1))
		wd = CINT(aSplitDate(2))
		DO
			nJulian = jdateFromWeeklyGregorian( wn, wd, mm, yy )
			'Response.Write "wn=" & wn & ", wd=" & wd & ", mm=" & mm & ", julian = " & nJulian & "<br>" & vbCRLF
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB
		
	PRIVATE SUB dateYearly_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("yearly")
		CASE "dayn"
			dateYearlyDayN_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "wday"
			dateYearlyWDay_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB
	
	
	PRIVATE FUNCTION dateRoshHashanah_calc( yy )
		dateRoshHashanah_calc = roshhashanah_conway( yy )
	END FUNCTION

	PRIVATE SUB dateRoshHashanah_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateRoshHashanah_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION datePassover_calc( yy )
		datePassover_calc = passover_conway( yy )
	END FUNCTION

	PRIVATE SUB datePassover_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = datePassover_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION dateOrthodoxEaster_calc( yy )
		DIM	d, m
		easter_mallen d, m, yy
		dateOrthodoxEaster_calc = jdateFromGregorian( d, m, yy )
	END FUNCTION

	PRIVATE SUB dateOrthodoxEaster_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateOrthodoxEaster_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB


	PRIVATE FUNCTION dateEaster_calc( yy )
		DIM	wd
		DIM	nJulian
		
		nJulian = paschal_moon( yy )
		wd = weekdayFromJDate( nJulian )
		wd = 1 - wd
		IF wd <= 0 THEN wd = wd + 7
		dateEaster_calc = nJulian + wd
	END FUNCTION
	
	PRIVATE SUB dateEaster_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = dateEaster_calc( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB

	PRIVATE SUB datePaschal_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	nJulian
		DIM	yy
		DIM	d, m, y

		gregorianFromJDate d, m, yy, nEarly
		DO
			nJulian = paschal_moon( yy )
			IF nLate < nJulian THEN EXIT DO
			IF nEarly <= nJulian THEN appendList nCount, aDates, nJulian
			yy = yy + 1
		LOOP
	END SUB
	
	PRIVATE SUB dateMoon_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
	END SUB


	PRIVATE SUB dateKeyword_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.text
		CASE "easter"
			dateEaster_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "orthodox easter"
			dateOrthodoxEaster_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "rosh hashanah"
			dateRoshHashanah_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "passover"
			datePassover_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "paschal"
			datePaschal_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "moon full", "moon new"
			dateMoon_calcDateList nCount, aDates, oDate, nEarly, nLate
		END SELECT
	END SUB
	

	PRIVATE SUB dateSeason_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		DIM	dBaseTime
		DIM	dTime
		DIM	nJulian
		SELECT CASE oDate.text
		CASE "0"
			dBaseTime = jtimeFromGregorian( 20, 3, 2000, 7, 35, 0 )
		CASE "1"
			dBaseTime = jtimeFromGregorian( 21, 6, 2000, 1, 48, 0 )
		CASE "2"
			dBaseTime = jtimeFromGregorian( 22, 9, 2000, 17, 27, 0 )
		CASE "3"
			dBaseTime = jtimeFromGregorian( 21, 12, 2000, 13, 37, 0 )
		CASE ELSE
			dBaseTime = 0
		END SELECT

		IF 0 < dBaseTime THEN
			nJulian = nEarly - FIX(dBaseTime)
			nJulian = FIX(nJulian / kTropicalYear)
			dTime = dBaseTime + (kTropicalYear * nJulian)
			IF nEarly < FIX(dTime) THEN dTime = dTime - kTropicalYear
			DO
				nJulian = FIX(dTime)
				IF nLate < nJulian THEN EXIT DO
				IF nEarly <= nJulian THEN appendList nCount, aDates, dTime
				dTime = dTime + kTropicalYear
			LOOP
		END IF
	END SUB
	
	PRIVATE SUB date_calcDateList( nCount, aDates, oDate, ByVal nEarly, ByVal nLate )
		SELECT CASE oDate.getAttribute("type")
		CASE "single"
			dateSingle_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "monthly"
			dateMonthly_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "yearly"
			dateYearly_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "keyword"
			dateKeyword_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "season"
			dateSeason_calcDateList nCount, aDates, oDate, nEarly, nLate
		CASE "julian"
		CASE ELSE
		END SELECT
	END SUB
	
	PRIVATE FUNCTION offset_calcEarly( nEarly, aOffset() )
		DIM	i
		DIM	nOffset
		
		nOffset = 0
		FOR i = 0 TO UBOUND(aOffset) STEP 2
			SELECT CASE aOffset(i)
			CASE "+", "++", "-", "--"
				nOffset = nOffset + 7
			CASE "#"
				nOffset = nOffset + ABS(CINT(aOffset(i+1)))
			CASE ELSE
			END SELECT
		NEXT 'i
		offset_calcEarly = nEarly - nOffset
	END FUNCTION

	PRIVATE FUNCTION offset_calcLate( nLate, aOffset() )
		DIM	i
		DIM	nOffset
		
		nOffset = 0
		FOR i = 0 TO UBOUND(aOffset) STEP 2
			SELECT CASE aOffset(i)
			CASE "+", "++", "-", "--"
				nOffset = nOffset + 7
			CASE "#"
				nOffset = nOffset - CINT(aOffset(i+1))
			CASE ELSE
			END SELECT
		NEXT 'i
		offset_calcLate = nLate + nOffset
	END FUNCTION
	
	PRIVATE FUNCTION offset_isHoliday( nDate )
		' NEEDS_WORK
		offset_isHoliday = FALSE
	END FUNCTION
	
	PRIVATE FUNCTION offset_evalCondition( nDate, sKwd )
		DIM	wd
		
		offset_evalCondition = FALSE
		wd = weekdayFromJDate( nDate )
		SELECT CASE sKwd
		CASE "weekday"
			IF 1 < wd  AND  wd < 7 THEN offset_evalCondition = TRUE
		CASE "weekend"
			IF wd < 2  OR  6 < wd THEN offset_evalCondition = TRUE
		CASE "workday"
			IF 1 < wd  AND  wd < 7  AND NOT offset_isHoliday( nDate ) THEN
				offset_evalCondition = TRUE
			END IF
		CASE "offday"
			IF wd < 2  OR  6 < wd  OR offset_isHoliday( nDate ) THEN
				offset_evalCondition = TRUE
			END IF
		CASE "holiday"
			IF offset_isHoliday( nDate ) THEN offset_evalCondition = TRUE
		CASE ELSE ' weekday number
			IF IsNumeric(sKwd) THEN
				IF wd = CINT(sKwd) THEN
					offset_evalCondition = TRUE
				END IF
			END IF
		END SELECT
	END FUNCTION
	
	PRIVATE FUNCTION offset_evalInc( nDate, sKwd, sValue )
		DIM	wd
		DIM	n
		
		wd = weekdayFromJDate( nDate )
		SELECT CASE sValue
		CASE "weekday", "workday"
			IF wd < 2  OR  6 < wd THEN
				SELECT CASE sKwd
				CASE "+", "++"
					IF wd < 2 THEN
						wd = 1
					ELSEIF 6 < wd THEN
						wd = 2
					ELSE
						wd = 0
					END IF
				CASE "-", "--"
					IF wd < 2 THEN
						wd = -2
					ELSEIF 6 < wd THEN
						wd = -1
					ELSE
						wd = 0
					END IF
				CASE "~"
					IF wd < 2 THEN
						wd = 1
					ELSEIF 6 < wd THEN
						wd = -1
					ELSE
						wd = 0
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		CASE "weekend", "offday"
			IF 1 < wd  AND  wd < 7 THEN
				SELECT CASE sKwd
				CASE "+", "++"
					wd = 7 - wd
				CASE "-", "--"
					wd = 1 - wd
				CASE "~"
					IF 3 < wd THEN
						wd = 7 - wd
					ELSE
						wd = 1 - wd
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		CASE "holiday"
			wd = 0
		CASE "cancel"
			wd = -32000
		CASE ELSE
			IF IsNumeric(sValue) THEN
				n = CINT(sValue)
				wd = n - weekdayFromJDate( nDate )
				SELECT CASE sKwd
				CASE "+"
					IF wd < 0 THEN wd = wd + 7
				CASE "++"
					IF wd <= 0 THEN wd = wd + 7
				CASE "-"
					IF 0 < wd THEN wd = wd - 7
				CASE "--"
					IF 0 <= wd THEN wd = wd - 7
				CASE "~"
					IF wd <= -3 THEN
						wd = wd + 7
					ELSEIF 3 < wd THEN
						wd = wd - 7
					END IF
				CASE ELSE
					wd = 0
				END SELECT
			ELSE
				wd = 0
			END IF
		END SELECT
		offset_evalInc = nDate + wd
	END FUNCTION
	
	PRIVATE FUNCTION offset_eval( nDate, aOffset )
		DIM	i
		
		i = 0
		DO WHILE i < UBOUND(aOffset)
			SELECT CASE aOffset(i)
			CASE "?"
				IF NOT offset_evalCondition( nDate, aOffset(i+1) ) THEN i = i + 2
			CASE "+", "++", "-", "--", "~"
				nDate = offset_evalInc( nDate, aOffset(i), aOffset(i+1) )
			CASE "#"
				nDate = nDate + CINT(aOffset(i+1))
			END SELECT
			i = i + 2
		LOOP
		offset_eval = nDate
	END FUNCTION
	
	PRIVATE SUB date_offset_calcDateList( nCount, aDates, oDate, oOffset, ByVal nEarly, ByVal nLate )
		DIM	sOffset
		DIM aList()
		DIM n, i
		DIM	nOfsEarly, nOfsLate
		DIM	nJulian
		
		sOffset = oOffset.value
		aSplitTemp = SPLIT( sOffset, " ", -1, vbTextCompare )
		nOfsEarly = offset_calcEarly( nEarly, aSplitTemp )
		nOfsLate = offset_calcLate( nLate, aSplitTemp )
		
		REDIM aList(10)
		n = 0
		date_calcDateList n, aList, oDate, nOfsEarly, nOfsLate
		IF 0 < n THEN
			FOR i = 0 TO n
				nJulian = offset_eval( aList(i), aSplitTemp )
				IF nEarly <= nJulian  AND  nJulian <= nLate THEN appendList nCount, aDates, nJulian
			NEXT 'i
		END IF
	END SUB
	
	PRIVATE SUB date_duration( nCount, aDates, ByVal nDur, ByVal nEarly, ByVal nLate )
		DIM	a()
		DIM n, i, j, k
		DIM d
		
		d = nDur
		IF d < 1 THEN d = 1
		REDIM a(nCount)
		FOR i = 0 TO nCount-1
			a(i) = aDates(i)
		NEXT 'i
		n = nCount-1
		nCount = 0
		FOR i = 0 TO n
			j = a(i)
			FOR k = 1 TO d
				IF nEarly <= j  AND  j <= nLate THEN appendList nCount, aDates, j
				j = j + 1
			NEXT 'k
		NEXT 'i
	END SUB


	SUB getDates( pCalendar )
	
		DIM	oEvents
		DIM	oEvent
		DIM	oCategories
		DIM	oCategory
		DIM	oDates
		DIM	oDate
		DIM	oItem
		DIM	bProcess
		
		DIM	sID
		DIM	sStyle
		DIM	sSubject
		DIM	sBody
		DIM	sLocation
		DIM	sSingle
		DIM	nJulian
		DIM	nDuration
		DIM	nPending
		DIM	nEarly
		DIM	nLate
		DIM i
		
		DIM	aDates()
		DIM	nDatesCount
		
		REDIM aDates(20)
		
		
		m_oXML.async = false
		m_oXML.load( m_sFilename )

		IF m_oXML.parseError.errorCode <> 0 THEN
			Response.Write "Error in File: " & Server.HTMLEncode(m_sFilename) & "<br>" & vbCRLF
			Response.Write "Error Code: " & m_oXML.parseError.errorCode & "<br>" & vbCRLF
			Response.Write "Error Reason: " & m_oXML.parseError.reason & "<br>" & vbCRLF
			Response.Write "Error Line: " & m_oXML.parseError.line & "<br>" & vbCRLF
		END IF
		
		SET oEvents = m_oXML.selectNodes("reminders/event" & m_sCategoryFilter)
		FOR EACH oEvent IN oEvents
			SET oItem = oEvent.attributes.getNamedItem("id")
			IF NOT Nothing IS oItem THEN
				sID = oItem.value
			ELSE
				sID = ""
			END IF

			SET oItem = oEvent.selectSingleNode("subject")
			sSubject = oItem.text
			SET oItem = Nothing
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("location")
			IF NOT oItem IS Nothing THEN
				sLocation = oItem.text
			ELSE
				sLocation = ""
			END IF
			ON Error GOTO 0
			sStyle = ""
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("style")
			IF NOT oItem IS Nothing THEN sStyle = oItem.text
			ON Error GOTO 0
			sBody = ""
			ON Error RESUME Next
			SET oItem = oEvent.selectSingleNode("body")
			IF NOT oItem IS Nothing THEN sBody = oItem.text
			ON Error GOTO 0
			SET oDates = oEvent.getElementsByTagName("date")
			nDatesCount = 0
			FOR EACH oDate IN oDates
				nDuration = 0
				nEarly = m_nDateBegin
				nLate = m_nDateEnd
				IF 0 < m_nPending THEN
					nPending = 0
					SET oItem = oDate.attributes.getNamedItem("pending")
					IF NOT( Nothing IS oItem ) THEN
						nPending = CLNG(oItem.value)
						nLate = nEarly + nPending
					END IF
				END IF
				SET oItem = oDate.attributes.getNamedItem("duration")
				IF NOT( Nothing IS oItem ) THEN
					nDuration = CLNG(oItem.value)
					nEarly = nEarly - nDuration
				END IF
				SET oItem = oDate.attributes.getNamedItem("offset")
				IF Nothing IS oItem THEN
					date_calcDateList nDatesCount, aDates, oDate, nEarly, nLate
				ELSE
					date_offset_calcDateList nDatesCount, aDates, oDate, oItem, nEarly, nLate
				END IF
				IF 1 < nDuration THEN
					date_duration nDatesCount, aDates, nDuration, nEarly, nLate
					nEarly = nEarly + nDuration
				END IF
				SET oCategories = oEvent.selectNodes("category")

				IF 0 < nDatesCount THEN
					FOR i = 0 TO nDatesCount-1
						nJulian = FIX(aDates(i))
						IF nEarly <= nJulian  AND  nJulian <= nLate THEN
							IF 0 < LEN(sID) THEN pCalendar.id = sID
							pCalendar.juliandate = aDates(i)
							pCalendar.subject = sSubject
							IF 0 < LEN(sBody) THEN pCalendar.body = sBody
							IF 0 < LEN(sStyle) THEN pCalendar.style = sStyle
							FOR EACH oCategory IN oCategories
								pCalendar.category = oCategory.text
							NEXT 'oCategory
							pCalendar.outputMessage
						END IF
					NEXT 'i
					nDatesCount = 0
				END IF

			NEXT 'oDate
		NEXT 'oEvent
		
	END SUB
	
END CLASS 'CCalendarFile



'===========================================================



%>