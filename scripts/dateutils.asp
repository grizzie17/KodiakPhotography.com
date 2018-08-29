<%
'---------------------------------------------------------------------
'
'            Copyright 1986 .. 2004 Bear Consulting Group
'                          All Rights Reserved
'
'    This software-file/document, in whole or in part, including	
'    the structures and the procedures described herein, may not	
'    be provided or otherwise made available without prior written
'    authorization.  In case of authorized or unauthorized
'    publication or duplication, copyright is claimed.
'
'---------------------------------------------------------------------


SUB gregorianFromJDate( d, m, y, jd )

	DIM	mm, yy
	DIM	l, n

	l = jd + 68569
	n = ( 4 * l ) \ 146097
	l = l - ((146097 * n + 3) \ 4)
	yy = (4000 * (l+1)) \ 1461001
	l = l - (((1461 * yy) \ 4) - 31)
	mm = (80 * l) \ 2447

	d = l - (( 2447 * mm ) \ 80)
	l = mm \ 11
	m = mm + ( 2 - 12 * l )

	y = yy + ( 100 * ( n - 49 ) + l )

END SUB




FUNCTION jdateFromGregorian( d, m, y )

	DIM	jmon
	DIM	jdate
	DIM	dd, mm, yy

	dd = CLNG(d)
	mm = CLNG(m)
	yy = CLNG(y)

	IF dd < 0 THEN
		dd = dd + 1
		mm = mm + 1
		IF 12 < mm THEN
			mm = mm - 12
			yy = yy + 1
		END IF
	END IF

	jmon = INT((mm - 14) \ 12)

	jdate = dd - 32075 + INT((1461 * (yy + 4800 + jmon)) \ 4) _
			+ INT((367 * (mm - 2 - jmon * 12)) \ 12) _
			- INT((3 * INT((yy + 4900 + jmon) \ 100)) \ 4)
	jdateFromGregorian = jdate

END FUNCTION


FUNCTION jtimeFromGregorian( d, m, y, hr, mn, sc )
	DIM	jdate
	DIM	nSec
	
	jdate = jdateFromGregorian( d, m, y )
	nSec = sc + mn*60 + hr*3600
	jtimeFromGregorian = CDBL(jdate) + (CDBL(nSec) / CDBL(3600*24))
	
END FUNCTION


' 1=sunday
FUNCTION weekdayFromJDate( jd )

	weekdayFromJDate = (jd - 6) MOD 7 + 1

END FUNCTION


FUNCTION jdateFromWeeklyGregorian( wn, wday, m, y )

	DIM	jdate
	DIM	n
	DIM	weeknum, wd, mm, yy
	
	weeknum = wn
	wd = wday
	mm = m
	yy = y

	IF weeknum < 0 THEN
		mm = mm + 1
		IF 12 < mm THEN
			mm = 1
			yy = yy + 1
		END IF
		weeknum = weeknum + 1
	END IF

	'
	'	first let's come up with the julian date of the first of the
	'	month for the indicated month.
	'
	jdate = jdateFromGregorian( 1, mm, yy )
	n = weekdayFromJDate( jdate )
	n = wd - n
	IF n < 0 THEN n = n + 7
	jdate = jdate + n + (( weeknum - 1 ) * 7)
	
	'
	'	now that we have a julian date let's make sure that it really
	'	falls in the indicated month.
	'
	IF 0 < weeknum THEN
		mm = mm + 1
		IF 12 < mm THEN
			mm = 1
			yy = yy + 1
		END IF
		IF jdateFromGregorian( 1, mm, yy ) <= jdate THEN jdate = 0	' problem
	ELSE
		mm = mm - 1
		IF mm < 1 THEN
			mm = 12
			yy = yy - 1
		END IF
		IF jdate < jdateFromGregorian( 1, mm, yy ) THEN jdate = 0	' problem
	END IF

	jdateFromWeeklyGregorian = jdate

END FUNCTION


FUNCTION diffWeekdaysJDates( jdFrom , jdTo, bInclEnd )
	DIM	nDiff
	DIM	nOffset
	DIM	nFromWD
	DIM	nToWD
	DIM	nNextSat
	DIM	nNextSun
	
	nDiff = jdTo - jdFrom
	IF bInclEnd THEN
		nDiff = nDiff + 1
		nOffset = 0
	ELSE
		nOffset = 1
	END IF
	nFromWD = weekdayFromJDate( jdFrom )
	nToWD = weekdayFromJDate( jdTo )
	nNextSat = jdFrom + (7 - nFromWD)
	nNextSun = jdFrom - nFromWD + 1
	IF 1 < nFromWD THEN nNextSun = nNextSun + 7
	diffWeekdaysJDates = nDiff
	IF nNextSat <= jdTo-nOffset THEN
		diffWeekdaysJDates = diffWeekdaysJDates - ((jdTo - nNextSat - nOffset) \ 7 + 1)
	END IF
	IF nNextSun <= jdTo-nOffset THEN
		diffWeekdaysJDates = diffWeekdaysJDates - ((jdTo - nNextSun - nOffset) \ 7 + 1)
	END IF
END FUNCTION



DIM aPaschalTable(2,19)
aPaschalTable(0,0) = 4
aPaschalTable(1,0) = 14
aPaschalTable(0,1) = 4
aPaschalTable(1,1) = 3
aPaschalTable(0,2) = 3
aPaschalTable(1,2) = 23
aPaschalTable(0,3) = 4
aPaschalTable(1,3) = 11
aPaschalTable(0,4) = 3
aPaschalTable(1,4) = 31
aPaschalTable(0,5) = 4
aPaschalTable(1,5) = 18
aPaschalTable(0,6) = 4
aPaschalTable(1,6) = 8
aPaschalTable(0,7) = 3
aPaschalTable(1,7) = 28
aPaschalTable(0,8) = 4
aPaschalTable(1,8) = 16
aPaschalTable(0,9) = 4
aPaschalTable(1,9) = 5
aPaschalTable(0,10) = 3
aPaschalTable(1,10) = 25
aPaschalTable(0,11) = 4
aPaschalTable(1,11) = 13
aPaschalTable(0,12) = 4
aPaschalTable(1,12) = 2
aPaschalTable(0,13) = 3
aPaschalTable(1,13) = 22
aPaschalTable(0,14) = 4
aPaschalTable(1,14) = 10
aPaschalTable(0,15) = 3
aPaschalTable(1,15) = 30
aPaschalTable(0,16) = 4
aPaschalTable(1,16) = 17
aPaschalTable(0,17) = 4
aPaschalTable(1,17) = 7
aPaschalTable(0,18) = 3
aPaschalTable(1,18) = 27
FUNCTION paschal_moon( yy )
	DIM	gn	' golden number
	
	gn = yy MOD 19 + 1 - 1	'(-1) to make golden-number an index
	paschal_moon = jdateFromGregorian( aPaschalTable(1,gn), _
									aPaschalTable(0,gn), yy )

END FUNCTION



SUB easter_mallen( d, m, ByVal y )

	d = 0
	m = 0

	' Calculate Easter Sunday date
	Dim FirstDig, Remain19, temp	'intermediate results (all integers)
	Dim tA, tB, tC, tD, tE			'table A to E results (all integers)

	FirstDig = y \ 100				'first 2 digits of year (\ means integer division)
	Remain19 = y Mod 19				'remainder of year / 19

	' calculate PFM date
	tA = ((225 - 11 * Remain19) Mod 30) + 21

	' find the next Sunday
	tB = (tA - 19) Mod 7
	tC = (40 - FirstDig) Mod 7

	temp = y Mod 100
	tD = (temp + temp \ 4) Mod 7

	tE = ((20 - tB - tC - tD) Mod 7) + 1
	d = tA + tE

	'10 days were 'skipped' in the Gregorian calendar from 5-14 Oct 1582
	temp = 10
	'Only 1 in every 4 century years are leap years in the Gregorian
	'calendar (every century is a leap year in the Julian calendar)
	If 1600 < y Then temp = temp + FirstDig - 16 - ((FirstDig - 16) \ 4)
	d = d + temp

	' return the date
	If 61 < d Then
		d = d - 61
		m = 5       'for method 2, Easter Sunday can occur in May
	ElseIf 31 < d Then
		d = d - 31
		m = 4
	Else
		m = 3
	End If

END SUB


FUNCTION passover_conway( ByVal y )

	DIM	j
	
	j = roshhashanah_conway( y )
	
	DIM	dd,mm,yy
	
	gregorianFromJDate dd, mm, yy, j
	
	IF 10 = mm THEN dd = dd + 30
	
	passover_conway = jdateFromGregorian( 21, 3, yy ) + mm


END FUNCTION


FUNCTION roshhashanah_conway( ByVal y )

	DIM	g	' golden number
	
	g = y MOD 19 + 1
	
	DIM	n
	DIM	r
	DIM	x
	
	x = (12*g) MOD 19
	
	n = ( y\100 - y\400 - 2 ) _
			+	765433/492480 * x _
			+	(y MOD 4)/4 _
			-	(313 * y + 89081)/98496
	
	r = n - INT(n)
	
	DIM	j
	
	j = jdateFromGregorian( INT(n), 9, y )
	
	DIM	wd
	
	wd = weekdayFromJDate( j )	'1=sunday
	
	SELECT CASE wd
	CASE 1,4,6	'sun,wed,fri
		j = j + 1
	CASE 2		'mon
		IF 23269/25920 <= r  AND  11 < x THEN
			j = j + 1
		END IF
	CASE 3		'tue
		IF 1367/2160 <= r  AND 6 < x THEN
			j = j + 2
		END IF
	END SELECT
	
	roshhashanah_conway = j
	

END FUNCTION



PUBLIC aDateSplit
SUB parseLogDate( dd, mm, yy, sDate )

	aDateSplit = SPLIT(sDate,"-",3,vbTextCompare)
	IF 2 = UBound(aDateSplit) THEN
		yy = 0 + aDateSplit(0)
		mm = 0 + aDateSplit(1)
		dd = 0 + aDateSplit(2)
	ELSE
		yy = 0
		mm = 0
		dd = 0
	END IF

END SUB



SUB parseLogDateEx( dd, mm, yy, sDate )

	DIM	dDate

	IF 0 < Len(sDate) THEN
		CALL parseLogDate( dd, mm, yy, sDate )
	ELSE
		dDate = DATE
		dd = DAY( dDate )
		mm = MONTH( dDate )
		yy = YEAR( dDate )
	END IF

END SUB


SUB gregorianFromVBDate( dd, mm, yy, dDate )

	dd = DAY( dDate )
	mm = MONTH( dDate )
	yy = YEAR( dDate )

END SUB

'FUNCTION jdateFromVBDate( dDate )
'	DIM	dd, mm, yy
'	dd = DAY( dDate )
'	mm = MONTH( dDate )
'	yy = YEAR( dDate )
'	jdateFromVBDate = jdateFromGregorian( dd, mm, yy )
'END FUNCTION


FUNCTION jdateFromVBDate( dDate )
	DIM	dd, mm, yy
	CALL gregorianFromVBDate( dd, mm, yy, dDate )
	jdateFromVBDate = jdateFromGregorian( dd, mm, yy )
END FUNCTION







FUNCTION logDateFromDDMMYY( dd, mm, yy )
	DIM	sTemp

	sTemp = CSTR(yy) & " "
	IF mm < 10 THEN
		sTemp = sTemp & "0" & mm & " "
	ELSE
		sTemp = sTemp & mm & " "
	END IF
	IF dd < 10 THEN
		sTemp = sTemp & "0" & dd
	ELSE
		sTemp = sTemp & dd
	END IF
	logDateFromDDMMYY = sTemp
END FUNCTION


FUNCTION logDateFromJDate( jd )
	DIM	dd, mm, yy
	CALL gregorianFromJDate( dd, mm, yy, jd )
	logDateFromJDate = logDateFromDDMMYY( dd, mm, yy )
END FUNCTION


SUB parseLogTime( hh, mm, ss, sTime )

	aDateSplit = SPLIT(sTime,":",3,vbTextCompare)
	hh = 0 + aDateSplit(0)
	mm = 0 + aDateSplit(1)
	ss = 0 + aDateSplit(2)

END SUB


%>