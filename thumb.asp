<%@ Language=VBScript %>
<%
OPTION EXPLICIT

Response.Expires = 30


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\ado.inc"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_picture.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<%



DIM	sTable
sTable = Request.QueryString("file")

DIM	q_h
DIM	q_w
q_h = Request.QueryString("h")
q_w = Request.QueryString("w")

DIM	q_copy
q_copy = Request.QueryString("c")


IF NOT IsNum(q_h) THEN q_h = 75
IF NOT IsNum(q_w) THEN q_w = 75






FUNCTION IsNum( s )
	IsNum = FALSE
	IF "" <> CSTR(s) THEN
		IF ISNUMERIC(s) THEN IsNum = TRUE
	END IF
END FUNCTION



FUNCTION getMime( s )

	getMime = "*/*"
	DIM	sSuffix
	DIM	aSuffix
	DIM	aMime
	DIM	i
	DIM	oFile
	DIM	sFile
	DIM	sData
	
	sSuffix = LCASE( s )
	sFile = findFileUpTree("scripts\mime.txt")
	IF "" <> sFile THEN
		SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
		IF NOT oFile IS Nothing THEN
			sData = oFile.ReadAll
			oFile.Close
			SET oFile = Nothing
			aMime = SPLIT(sData, vbCRLF )
			DIM	sLine
			FOR EACH sLine IN aMime
				IF "" <> sLine THEN
					aSuffix = SPLIT(sLine,vbTAB)
					IF LBOUND(aSuffix) < UBOUND(aSuffix) THEN
						IF sSuffix = aSuffix(0) THEN
							getMime = aSuffix(1)
							EXIT FOR
						END IF
					END IF
				END IF
			NEXT
		END IF
	END IF

END FUNCTION

FUNCTION getContentType( s )

	DIM	i
	DIM	sSuffix
	i = INSTRREV( s, ".", -1, vbTextCompare )
	IF 0 < i THEN
		sSuffix = MID( s, i+1 )
		getContentType = getMime( sSuffix )
		IF "asp" = sSuffix  OR  "xslt" = sSuffix THEN
			getContentType = ""
		END IF
	ELSE
		getContentType = "image/gif"
	END IF
	
END FUNCTION




SUB publishFile( cType, sPath )
	DIM	oStream
	
	SET oStream = Server.CreateObject("ADODB.Stream")
	
	IF NOT oStream IS Nothing THEN

		oStream.Open
		oStream.Type = adTypeBinary
	
		oStream.LoadFromFile sPath
	
		Response.ContentType = cType
		Response.BinaryWrite oStream.Read
		
		oStream.Close
		SET oStream = Nothing
	
	END IF
END SUB


' assumes that canvas has already been loaded and cropped
SUB thumbResize( oCanvas, h, w )

	IF IsNum(q_h)  AND  IsNum(q_w) THEN
	
		DIM	wRatio
		DIM	hRatio
		DIM	fRatio
	
		wRatio = w / CDBL(q_w)
		hRatio = h / CDBL(q_h)
		IF hRatio < wRatio THEN
			fRatio = wRatio
		ELSE
			fRatio = hRatio
		END IF
		w = FIX(w / fRatio)
		h = FIX(h / fRatio)
		oCanvas.Width = w
		oCanvas.Height = h
		
	END IF

END SUB


FUNCTION makeCanvas( cType, sSrcPath )

	SET makeCanvas = Nothing
	
	DIM	oCanvas
	SET oCanvas = Nothing
	IF 0 < INSTR( ctype, "jpeg" ) THEN
		ON ERROR Resume Next
		SET oCanvas = Server.CreateObject("Persits.Jpeg")
		ON ERROR Goto 0
		IF NOT oCanvas IS Nothing THEN
			'Response.Write sSrcPath & "<br>"
			'Response.Flush
			oCanvas.Open sSrcPath
			SET makeCanvas = oCanvas
			SET oCanvas = Nothing
		END IF
	END IF
END FUNCTION


FUNCTION thumbMakePredefined( sCache, sFile, cType, sSrcPath )
	thumbMakePredefined = ""
	
	DIM	oCanvas
	SET oCanvas = makeCanvas( cType, sSrcPath )
	IF NOT oCanvas IS Nothing THEN
		DIM	h
		DIM	w
		
		h = oCanvas.OriginalHeight
		w = oCanvas.OriginalWidth
	
		thumbResize oCanvas, h, w
		
		thumbMakePredefined = cache_filepath( sCache, sFile )
		IF g_oFSO.FileExists( thumbMakePredefined ) THEN
			g_oFSO.DeleteFile thumbMakePredefined
		END IF
		IF "" <> thumbMakePredefined THEN
			oCanvas.Save thumbMakePredefined
		END IF
		
		SET oCanvas = Nothing
	ELSE
		thumbMakePredefined = sSrcPath
	END IF
	
END FUNCTION


FUNCTION thumbMake( sCache, sFile, cType, sSrcPath )
	thumbMake = ""
	
	DIM	oCanvas
	SET oCanvas = makeCanvas( cType, sSrcPath )
	
	IF NOT oCanvas IS Nothing THEN
		
		DIM	x, y
		DIM	h, w
		h = oCanvas.OriginalHeight
		w = oCanvas.OriginalWidth
		
		IF NOT bThumbDefined THEN
			IF h < w THEN	' landscape
				i = h \ 8	' define margin
				x = (w - (i*6)) \ 2
				y = i
				w = h - (i * 2)
				h = w
				oCanvas.crop x, y, x+w, y+h
			ELSE			' portrait
				i = w \ 8	' define margin
				y = (h - (i*6)) \ 4
				x = i
				h = w - (i * 2)
				w = h
				oCanvas.crop x, y, x+w, y+h
			END IF
		END IF
		
		thumbResize oCanvas, h, w
		
		
		thumbMake = cache_filepath( sCache, sFile )
		IF "" <> thumbMake THEN
			oCanvas.Save thumbMake
		END IF
		
		SET oCanvas = Nothing
	ELSE
		thumbMake = sSrcPath
	END IF
	'Response.write thumbMake & "<br>"

END FUNCTION


SUB publishThumbFile( cType, sPath, sTable, bThumbDefined )

	DIM	sCache
	IF IsNum(q_h) THEN
		DIM	i
		DIM	sTbl
		sTbl = REPLACE( sTable, "/", "\" )
		IF "\" <> LEFT( sTbl, 1 ) THEN sTbl = "\" & sTbl
		i = INSTRREV( sTbl, "\" )
		IF 0 < i THEN
			sCache = "thumb" & "\" & CSTR(q_h) & LEFT(sTbl,i-1)
			
			DIM	sFile
			sFile = MID(sTable,i+1)			
			
			DIM	sCachePath
			sCachePath = cache_checkFile( sCache, sFile, "d", 14, "m" )
			IF 0 < LEN(sCachePath) THEN
				publishFile cType, sCachePath
			ELSE
				IF bThumbDefined THEN
					sCachePath = thumbMakePredefined( sCache, sFile, cType, sPath )
				ELSE
					sCachePath = thumbMake( sCache, sFile, cType, sPath )
				END IF
				publishFile cType, sCachePath
			END IF
		ELSE
			publishFile cType, sPath
		END IF
	ELSE
		publishFile cType, sPath	
	END IF

END SUB





DIM	sName

DIM	sPicturePath
DIM	bThumbDefined
bThumbDefined = FALSE

DIM	i
DIM	s, t

IF "" <> sTable THEN

	IF ":" = MID(sTable,2,1) THEN
		sPicturePath = sTable
	ELSE
		sPicturePath = Server.MapPath(sTable)
	END IF
	IF NOT g_oFSO.FileExists( sPicturePath ) THEN
		sPicturePath = Server.MapPath("images/notfound.jpg")
	ELSE
		i = INSTRREV( sPicturePath, "\" )
		IF 0 < i THEN
			s = LEFT( sPicturePath, i )
			t = MID( sPicturePath, i+1 )
			s = s & "_thumbs\" & t
			IF g_oFSO.FileExists( s ) THEN
				sPicturePath = s
				bThumbDefined = TRUE
			END IF
		END IF
	END IF

ELSE
	sPicturePath = Server.MapPath("images/notfound.jpg")
END IF

DIM	cType
cType = getContentType( sPicturePath )
IF "" <> cType THEN

	publishThumbFile cType, sPicturePath, sTable, bThumbDefined



END IF

SET g_oFSO = Nothing


%>
