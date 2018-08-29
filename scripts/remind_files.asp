<%

DIM	dRemindLastModified


FUNCTION loadRemindFiles( nDateBegin, nDateEnd, sCategories, sHolidayCategories, bShowToday )

	DIM	sRemindFile
	DIM	oCalendar
	DIM	oCalendarFile
	
	DIM	aFileList()
	REDIM aFileList(5)
	DIM	aTempSplit
	DIM	sTemp
	DIM	dateTemp
	
	getRemindList aFileList
	
	
	SET loadRemindFiles = Nothing
	
	IF -1 < UBOUND(aFileList) THEN
	
		aTempSplit = SPLIT( aFileList(0), vbTAB )
		dRemindLastModified = CDATE( aTempSplit(2) )
		
		SET	oCalendar = new CCalendar
		
		IF bShowToday THEN
			oCalendar.juliandate = jdateFromVBDate( NOW )
			oCalendar.subject = "* T * O * D * A * Y *"
			oCalendar.style = "RmdToday"
			oCalendar.outputMessage
		END IF
		
		FOR EACH sTemp IN aFileList
			IF "" <> sTemp THEN
				aTempSplit = SPLIT( sTemp, vbTAB )
				dateTemp = CDATE(aTempSplit(2))
				IF dRemindLastModified < dateTemp THEN dRemindLastModified = dateTemp
				SET oCalendarFile = new CCalendarFile
				oCalendarFile.file = aTempSplit(1)
				oCalendarFile.datebegin = nDateBegin
				oCalendarFile.dateend = nDateEnd
				IF 0 < INSTR(aTempSplit(1),";categories") THEN
					IF 0 < LEN(sCategories) THEN oCalendarFile.categories = sCategories
				ELSEIF 0 < INSTR(aTempSplit(1),";holiday") THEN
					IF 0 < LEN(sHolidayCategories) THEN oCalendarFile.categories = sHolidayCategories
				END IF
				oCalendarFile.getDates( oCalendar )
			END IF
		NEXT 'sTemp
				
		SET loadRemindFiles = oCalendar
		SET oCalendar = Nothing
		SET oCalendarFile = Nothing

	END IF
END FUNCTION
%>