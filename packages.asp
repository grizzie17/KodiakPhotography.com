<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<%


DIM	aSubList()
REDIM aSubList(5)
DIM	nSubCount
nSubCount = 0
DIM	x_sPage
DIM	x_nIndex
DIM	x_sPageTitle
DIM	x_sFile


DIM	ifn
DIM	sFolderName
sFolderName = Request.ServerVariables("SCRIPT_NAME")
ifn = INSTRREV( sFolderName, "/" )
IF 0 < ifn THEN
	sFolderName = MID( sFolderName, ifn+1 )
END IF
ifn = INSTRREV( sFolderName, "." )
IF 0 < ifn THEN
	sFolderName = LEFT( sFolderName, ifn-1 )
END IF


	g_sPage = Request("page")
	g_sUseFileNameSuffix = ".htm"
	buildFileListX nSubCount, aSubList, sFolderName, FALSE

	x_sFile = ""
	x_nIndex = 0
	FOR nLen = 0 TO nSubCount-1
		aFileSplit = SPLIT( aSubList(nLen), vbTAB, -1, vbTextCompare )
		IF LCASE(g_sPage) = LCASE(aFileSplit(kFI_Name)) THEN
			x_sPageTitle = aFileSplit(kFI_Title)
			x_sFile = aFileSplit(kFI_Name)
			x_nIndex = nLen
			EXIT FOR
		END IF
	NEXT 'nLen
	IF "" = x_sFile THEN
		aFileSplit = SPLIT( aSubList(0), vbTAB, -1, vbTextCompare )
		x_sFile = aFileSplit(kFI_Name)
		x_sPageTitle = aFileSplit(kFI_Title)
	END IF






%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
IF TRUE THEN
	Response.Write "<" & "title" & ">"
	IF "" <> x_sPageTitle THEN Response.Write Server.HTMLEncode(x_sPageTitle) & " - "
	IF "" <> g_navtab_sTitle THEN Response.Write Server.HTMLEncode(g_navtab_sTitle) & " - "
	Response.Write Server.HTMLEncode(g_sSiteName)
	Response.Write "<" & "/title" & ">" & vbCRLF
ELSE
%>
<title>Subfolder List</title>
<%
END IF
%>
<meta name="keywords" content="photography, photographer, huntsville, alabama, wedding, portrait, scenery">
<meta name="navigate" content="tab">
<meta name="navtitle" content="Packages">
<meta name="sortname" content="zpackages">
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<%
DIM	sStylePath
sStylePath = g_sTheme & sFolderName & ".css"
IF g_oFSO.FileExists(Server.MapPath(sStylePath)) THEN
%>
<link rel="stylesheet" href="<%=sStylePath %>" type="text/css">
<%
END IF
IF g_oFSO.FileExists(Server.MapPath(sFolderName & "/style.css")) THEN
%>
<link rel="stylesheet" href="<%=sFolderName%>/style.css" type="text/css">
<%
END IF
%>
</head>

<body>


<!--#include file="scripts\page_begin.asp"-->
<!--#include file="scripts\include_navigation.asp"-->
<%


CLASS CTabFormatGray

	PROPERTY GET colorBackground()
		colorBackground = "#999999"
	END PROPERTY
	
	PROPERTY GET colorTab()
		colorTab = "#CCCCCC"
	END PROPERTY
	
	PROPERTY GET colorTabSelected()
		colorTabSelected = "#FFFFFF"
	END PROPERTY
	
	PROPERTY GET classTabGroup()
		classTabGroup = "SubNavigation"
	END PROPERTY
	
	PROPERTY GET classTabGroupInverted()
		classTabGroupInverted = "SubNavigationInverted"
	END PROPERTY

	PROPERTY GET classTabGroupVert()
		classTabGroupVert = "TabGroupVert"
	END PROPERTY
	

	PROPERTY GET classTab()
		classTab = "shoppingtab"
	END PROPERTY
	
	PROPERTY GET classSelected()
		classSelected = "SelectedTab"
	END PROPERTY
	
	PROPERTY GET alignTabHoriz()
		alignTabHoriz = "center"
	END PROPERTY
	
	PROPERTY GET imageTL()
		imageTL = "images/pie_tl_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageTR()
		imageTR = "images/pie_tr_gray.gif"
	END PROPERTY
	
	PROPERTY GET imageBL()
		imageBL = "images/pie_bl_gray.gif"
	END PROPERTY

	PROPERTY GET imageBR()
		imageBR = "images/pie_br_gray.gif"
	END PROPERTY
	
END CLASS


CLASS CTabMemberData

	PRIVATE m_aData
	PRIVATE m_aSplit
	PRIVATE m_i
	PRIVATE m_sURL

	PRIVATE SUB Class_Initialize
		m_sURL = ""
		m_i = 0
	END SUB

	
	PROPERTY GET RecordCount()
		RecordCount = UBOUND(m_aData) + 1
	END PROPERTY
	
	PROPERTY GET EOF()
		IF m_i <= UBOUND(m_aData) THEN
			EOF = FALSE
		ELSE
			EOF = TRUE
		END IF
	END PROPERTY
	
	SUB MoveFirst()
		m_i = 0
		privParse
	END SUB
	
	SUB MoveNext()
		m_i = m_i + 1
		privParse
	END SUB
	
	FUNCTION IsCurrent( x )
		IF m_i = x THEN
			IsCurrent = TRUE
		ELSE
			IsCurrent = FALSE
		END IF
	END FUNCTION
	
	PROPERTY GET HREF()
		HREF = m_sURL & m_aSplit(kFI_Name)
	END PROPERTY
	
	PROPERTY GET TabLabel()
		TabLabel = m_aSplit(kFI_Title)
	END PROPERTY
	
	'----------------
	
	PRIVATE SUB privParse()
		IF NOT EOF() THEN
			m_aSplit = SPLIT( m_aData(m_i), vbTAB, -1, vbTextCompare )
		END IF
	END SUB
	
	PROPERTY LET Data( a )
		m_aData = a
	END PROPERTY
	
	PROPERTY LET URL( s )
		m_sURL = s
	END PROPERTY
	
END CLASS


	DIM	oTabGen
	DIM	oTabData
	DIM	oTabFormat
	
	
	SET oTabGen = NEW CTabGenerate
	SET oTabData = NEW CTabMemberData
	SET oTabFormat = NEW CTabFormatGray
	
	
	oTabData.Data = aSubList
	oTabData.URL = sFolderName & ".asp?page="
	
	SET oTabGen.TabData = oTabData
	SET oTabGen.TabFormat = oTabFormat
	oTabGen.MaxCols = 1
	oTabGen.Current = x_nIndex
	oTabGen.makeTabs
	
%>
<div id="bodycontent">
<%
	includeBody sFolderName & "\" & x_sFile
%>
</div>
<%


	'oTabGen.makeTabsInverted
	
	SET oTabFormat = Nothing
	SET oTabData = Nothing
	SET oTabGen = Nothing





%>


<!--#include file="scripts\page_end.asp"-->
<!--webbot bot="Include" TAG="BODY" U-Include="_private/byline.htm" startspan -->

<script language="JavaScript">
<!--

function makeByLine()
{
	document.write( '<' + 'script language="JavaScript" src="http://' );
	if ( "localhost" == location.hostname )
	{
		document.writeln( 'localhost/BearConsultingGroup' );
	}
	else
	{
		document.writeln( 'BearConsultingGroup.com' );
	}
	document.writeln( '/designby_small.js"></' + 'script>' );
}
makeByLine()

//-->
</script>
<!--script language="JavaScript" src="http://BearConsultingGroup.com/designbyadvert.js"></script-->

<!--webbot bot="Include" i-checksum="20551" endspan -->

</body>

</html>
<%
SET g_oFSO = Nothing
%>