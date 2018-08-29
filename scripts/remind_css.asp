<%@ Language=VBScript %>
<%
OPTION EXPLICIT

Response.Expires = 30

DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")



%>
<!--#include file="remind_utils.asp"-->
<%


DIM	sFile

sFile = findRemindFile( "remind.css" )
IF "" <> sFile THEN

	DIM	oFile
	DIM	sText
	
	SET oFile = g_oFSO.OpenTextFile( sFile, 1 )
	IF NOT oFile IS Nothing THEN
	
		sText = oFile.ReadAll
		
		oFile.Close
		SET oFile = Nothing
		
		Response.ContentType = "text/css"
		Response.Write sText
		IF Response.Buffer THEN Response.Flush
	
	END IF

END IF

SET g_oFSO = Nothing

%>