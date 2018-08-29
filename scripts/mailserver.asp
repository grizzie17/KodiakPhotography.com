<%
FUNCTION mailServer()
	DIM	sServer
	DIM	i
	sServer = Request.ServerVariables("LOCAL_ADDR")
	i = INSTRREV( sServer, ".", -1, vbTextCompare )
	IF 0 < i THEN
		mailServer = LEFT(sServer,i) & CSTR( CINT(MID(sServer,i+1)) + 4 )
	ELSE
		mailServer = sServer
	END IF
	sServer = Request.ServerVariables("HTTP_HOST")
	IF "localhost" = LCASE(sServer) THEN
		mailServer = "localhost"
	ELSE
		mailServer = Request.ServerVariables("LOCAL_ADDR")
		'mailServer = "mail.grizzlyweb.com"
	END IF
END FUNCTION

%>