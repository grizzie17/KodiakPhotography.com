<!--#include file="sortutil.asp"-->
<%



FUNCTION getBasePart( sPath )

	DIM	sPart
	DIM	i
	
	sPart = sPath
	i = INSTRREV(sPart,"\")
	IF 0 < i THEN
		sPart = MID(sPart,i+1)
	END IF
	
	i = INSTRREV(sPart,"." )
	IF 0 < i THEN
		sPart = LEFT(sPart,i-1)
	END IF
	
	getBasePart = sPart

END FUNCTION





DIM	g_sNewletterFolder
DIM	g_sRootFolder


FUNCTION pagePicture( sLabel )

	IF g_sNewletterFolder <> "" THEN
		pagePicture = g_sNewletterFolder & "\" & sLabel
		pagePicture = "picture.asp?file=" & Server.URLEncode(pagePicture)
	'	pagePicture = MID(pagePicture, LEN(g_sRootFolder)+1)
	'	pagePicture = REPLACE(pagePicture, "\", "/")
	END IF

END FUNCTION

g_htmlFormat_pictureFunc = "pagePicture"




SUB getAllNewsletters( aLetters() )

	DIM	sFolder
	sFolder = findFolder( "cgi-bin\newsletters" )
	g_sNewletterFolder = sFolder
	g_sRootFolder = Server.MapPath( "/" )
	
	IF "" <> sFolder THEN
		DIM	oFolder
		DIM	oFile
		DIM	sFile
		DIM	i
		
		SET oFolder = g_oFSO.GetFolder( sFolder )
		
		i = 0
		REDIM aLetters(5)
		FOR EACH oFile IN oFolder.Files
			sFile = LCASE(oFile.Name)
			IF 0 < INSTR(1, sFile, ".txt", vbTextCompare ) THEN
				IF UBOUND(aLetters) <= i THEN
					REDIM PRESERVE aLetters(i+5)
				END IF
				aLetters(i) = LCASE(oFile.Path)
				i = i + 1
			END IF
		NEXT 'oFile
		
		IF 0 < i THEN
			REDIM PRESERVE aLetters(i-1)
			sortDescend aLetters, 0, i-1
		END IF
		
		SET oFolder = Nothing
	END IF

END SUB




%>
