<%@	Language=VBScript
	EnableSessionState=True %>
<%
OPTION EXPLICIT


DIM	g_oFSO
SET g_oFSO = Server.CreateObject("Scripting.FileSystemObject")

%>
<!--#include file="gallery_library.asp"-->
<%



SUB getGalleryFilenames()

	getPictureFilenames g_aGalleryFilenames, g_nGalleryIndexCount, "Gallery\Weddings", FALSE
	
END SUB

getGalleryFilenames


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Weddings - Gallery - Kodiak Photography</title>
<meta name="keywords" content="photography, photographer, huntsville, alabama, wedding, portrait, scenery">
<meta name="navigate" content="tab">
<meta name="sortname" content="aa">
<meta name="navtitle" content="Weddings">
<!--#include file="scripts\favicon.asp"-->
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<%
makeJSImageList "Gallery", g_aGalleryFilenames, g_nGalleryIndexCount
makeGalleryJavascript
makePanelStyles
%>


</head>

<body>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="scripts\include_navigation.asp"-->
<%
makeGalleryBody
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