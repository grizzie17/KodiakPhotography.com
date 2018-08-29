<%

SUB makeFavIconLink()

	DIM	s
	DIM	i
	DIM	oFolder
	DIM	sIconPath
	
	SET oFolder = g_oFSO.GetFolder(Server.MapPath("."))

	Response.Write "<link rel=""shortcut icon"" type=""image/ico"" href=""http://"
	
	DIM sHost
	sHost = Request.ServerVariables("HTTP_HOST")
	
	Response.Write sHost
	
	s = Request.ServerVariables("URL")
	i = INSTRREV( s, "/" )
	IF 0 < i THEN
		s = LEFT( s, i-1 )
	ELSE
		s = ""
	END IF
	Response.Write s
	Response.Write "/favicon.ico"">" & vbCRLF

	sIconPath = g_oFSO.BuildPath( oFolder.Path, "favicon144.png" )
	IF g_oFSO.FileExists( sIconPath ) THEN
		Response.Write "<meta name=""msapplication-TileImage"" content=""http://" & sHost & s & "/favicon144.png"" />" & vbCRLF
	END IF
	SET oFolder = Nothing


END SUB
makeFavIconLink

%>