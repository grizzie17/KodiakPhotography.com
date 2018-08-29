<!--#include file="ado.inc"-->
<!--#include file="findfiles.asp"-->
<%


FUNCTION dbSessionKey()
	DIM	i
	DIM	sSessionKey
	sSessionKey = "MDBFile" & Request.ServerVariables("PATH_INFO")
	i = INSTRREV(sSessionKey,"/")
	IF 0 < i THEN sSessionKey = LEFT(sSessionKey,i-1)
	dbSessionKey = sSessionKey
END FUNCTION


FUNCTION dbConnect()

	SET dbConnect = Nothing
	
	DIM	sSessionKey
	sSessionKey = dbSessionKey()
	
	DIM	sDBFile
	sDBFile = Session.Contents(sSessionKey)
	IF "" = sDBFile THEN
		sDBFile = findDBFile( "Products.mdb" )
	END IF
	
	DIM	oDC
	SET oDC = Nothing
	IF "" <> sDBFile THEN
		Session(sSessionKey) = sDBFile
		SET oDC = Server.CreateObject( "ADODB.Connection" )
		IF NOT Nothing IS oDC THEN
			DIM	sConn
			sConn = "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & sDBFile & ";"
			oDC.open sConn
			SET dbConnect = oDC
			SET oDC = Nothing
		END IF
	END IF

END FUNCTION



FUNCTION dbQueryRead( oDC, sQuery )

	DIM	oRS
	SET oRS = Server.CreateObject("ADODB.RecordSet")
	IF NOT Nothing IS oRS THEN
		With oRS
			.CursorLocation = 3 'adUseClient
			.CursorType     = adOpenStatic 'adOpenKeyset
			.LockType       = adLockReadOnly
			.Open sQuery, oDC
			Set .ActiveConnection = Nothing
		End With
	END IF
	'SET oRS = oDC.Execute( sQuery )
	SET dbQueryRead = oRS
	SET oRS = Nothing

END FUNCTION


FUNCTION dbQueryUpdate( oDC, sQuery )

	SET dbQueryUpdate = Nothing

	DIM	oRS
	SET oRS = Server.CreateObject("ADODB.RecordSet")
	IF NOT Nothing IS oRS THEN
		With oRS
			.CursorLocation = 3 'adUseClient
			.CursorType     = adOpenDynamic
			.LockType       = adLockOptimistic
			.Open sQuery, oDC
		End With
	END IF
	'SET oRS = oDC.Execute( sQuery )
	SET dbQueryUpdate = oRS
	SET oRS = Nothing

END FUNCTION


FUNCTION recString( o )
	IF ISNULL( o ) THEN
		recString = ""
	ELSEIF ISNULL( o.Value ) THEN
		recString = ""
	ELSE
		recString = TRIM(CSTR(o.Value))
	END IF
END FUNCTION


FUNCTION recNumber( o )
	IF ISNULL( o ) THEN
		recNumber = 0
	ELSEIF ISNULL( o.Value ) THEN
		recNumber = 0
	ELSEIF ISNUMERIC( o.Value ) THEN
		recNumber = CLNG(o.Value)
	ELSE
		recNumber = 0
	END IF
END FUNCTION


FUNCTION fieldString( s )
	IF "" = s THEN
		fieldString = NULL
	ELSE
		fieldString = s
	END IF
END FUNCTION


FUNCTION fieldBool( o )
	fieldBool = 0
	IF "ON" = UCASE(o) THEN
		fieldBool = -1
	END IF
END FUNCTION


FUNCTION fieldNumber( s )
	IF "" = CSTR(s) THEN
		fieldNumber = NULL
	ELSEIF ISNUMERIC( s ) THEN
		fieldNumber = s
	ELSE
		fieldNumber = 0
	END IF
END FUNCTION





DIM g_oFSO
SET g_oFSO = CreateObject( "Scripting.FileSystemObject" )




DIM g_DC
SET g_DC = dbConnect()




%>