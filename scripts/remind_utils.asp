<%

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


DIM	g_remind_sDataFolder
g_remind_sDataFolder = ""


FUNCTION findRemindFolder()
	DIM	sFolder
	
	IF "" = g_remind_sDataFolder THEN
		sFolder = findFolder( "database\remind" )
		IF "" = sFolder THEN
			sFolder = findFolder( "cgi-bin\remind" )
		END IF
		IF "" <> sFolder THEN
			findRemindFolder = sFolder
			g_remind_sDataFolder = sFolder
		ELSE
			findRemindFolder = ""
		END IF
	ELSE
		findRemindFolder = g_remind_sDataFolder
	END IF
END FUNCTION


FUNCTION getRemindSortname( sName )
	DIM	i
	i = INSTR(sName,";sort=")
	IF 0 < i THEN
		DIM	sTemp
		sTemp = MID(sName,i+6)
		i = INSTR(sTemp,";")
		IF 0 < i THEN
			sTemp = LEFT(sTemp,i-1)
		ELSE
			sTemp = LEFT(sTemp,LEN(sTemp)-4)
		END IF
		i = INSTR(sName,";")
		IF 0 < i THEN
			getRemindSortname = sTemp & LEFT(sName,i-1)
		ELSE
			getRemindSortname = sTemp & LEFT(sName,LEN(sName)-4)
		END IF
	ELSE
		getRemindSortname = sName
	END IF
END FUNCTION


SUB getRemindList( aList() )
	DIM	sFolder
	DIM	oFolder
	sFolder = findRemindFolder()
	IF "" <> sFolder THEN
		SET oFolder = g_oFSO.GetFolder( sFolder )
		IF NOT Nothing IS oFolder THEN
			DIM	nCount
			DIM	oFile
			DIM	sName
			DIM	sNameLC
			DIM	sFile
			DIM	sFileLC
			nCount = 0
			FOR EACH oFile IN oFolder.Files
				sName = oFile.Name
				sNameLC = LCASE(sName)
				IF ".xml" = RIGHT(sNameLC,4) THEN
					sFile = oFile.Path
					IF UBOUND(aList) <= nCount THEN
						REDIM PRESERVE aList(nCount+5)
					END IF
					aList(nCount) = getRemindSortname(sNameLC) & vbTAB & sFile & vbTAB & oFile.DateLastModified
					nCount = nCount + 1
				END IF
			NEXT 'sFile
			REDIM PRESERVE aList(nCount-1)
			sort aList, LBOUND(aList), UBOUND(aList)
		END IF
	END IF
END SUB


FUNCTION findRemindFile( sRemindName )
	DIM	sFolder
	DIM	sPath
	
	sFolder = findRemindFolder()
	IF 0 < LEN( sFolder ) THEN
		sPath = g_oFSO.BuildPath( sFolder, sRemindName )
		IF g_oFSO.FileExists( sPath ) THEN
			findRemindFile = sPath
			EXIT FUNCTION
		END IF
	END IF
	findRemindFile = ""
END FUNCTION


FUNCTION remindCSS()

	remindCSS = "scripts/remind_css.asp"

	'remindCSS = ""
	'DIM	sFile
	
	'sFile = findRemindFile( "remind.css" )
	'IF "" <> sFile THEN
	'	DIM	sRootPath
	'	DIM	sVirtPath
	'	DIM	sTemp
	'	sRootPath = Server.MapPath("/")
	'	sTemp = MID(sFile,LEN(sRootPath)+1)
	'	remindCSS = REPLACE(sTemp,"\","/",1,-1,vbTextCompare)
	'	IF "/" <> LEFT(remindCSS,1) THEN remindCSS = "/" & remindCSS
	'END IF
END FUNCTION


%>