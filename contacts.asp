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
<!--#include file="scripts/sortutil.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\include_gallery.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<%


DIM g_n
DIM g_sGalleryFilename
DIM g_aGalleryFilenames()
REDIM g_aGalleryFilenames(10)
DIM	g_sGalleryFolder
DIM g_nGalleryIndexCount
DIM g_dGalleryIndexTime
DIM g_sGalleryCookie
g_sGalleryFilename = ""
g_sGalleryFolder = ""
g_sGalleryCookie = ""
g_nGalleryIndexCount = 0
g_dGalleryIndexTime = CDATE("1/1/2000")
FOR g_n = LBOUND(g_aGalleryFilenames) TO UBOUND(g_aGalleryFilenames)
	g_aGalleryFilenames(g_n) = ""
NEXT


DIM	gPictureDate
gPictureDate = CDATE("1/1/2000")



SUB getPictureFilenames( BYREF aList(), BYREF nList, sBaseName, bRandomize )

	DIM	aData
	'REDIM aData(50)
	aData = Array("","")
	DIM	sCacheFilename
	sCacheFilename = REPLACE(sBaseName, "\", "-")
	sCacheFilename = REPLACE(sCacheFilename, "/", "-")
	sCacheFilename = sCacheFilename & "-contacts"
	DIM	sCacheData
	DIM	oFile
	SET oFile = cache_openTextFile( "shuffle", sCacheFilename & ".txt", "d", 7, "m" )
	IF NOT oFile IS Nothing THEN
		sCacheData = oFile.ReadAll
		oFile.Close
		SET oFile = Nothing
		aData = SPLIT(sCacheData, vbCRLF)
		nGalleryCount = UBOUND(aData) + 1	'account for zero index
		IF bRandomize THEN
			shuffle aData, 0, nGalleryCount-1
			shuffle aData, 0, nGalleryCount-1
		ELSE
			sort aData, 0, nGalleryCount-1
		END IF
	ELSE

		DIM	sFolder
		IF "" = g_sBaseDataFolder THEN
			sFolder = findAppDataFolder( sBaseName )
		ELSE
			sFolder = g_oFSO.BuildPath( g_sBaseDataFolder, sBaseName )
		END IF
		
		DIM	dTemp
		
		bGalleryRandomize = bRandomize
		
		readGalleryFolder sFolder, aData
		
		IF 0 < nGalleryCount THEN
			SET oFile = cache_makeFile( "shuffle", sCacheFilename & ".txt" )
			IF NOT oFile IS Nothing THEN
				sCacheData = JOIN(aData, vbCRLF)
				oFile.Write sCacheData
				oFile.Close
				SET oFile = Nothing
			END IF
		END IF

	END IF

	
	IF 0 < nGalleryCount THEN

	
		nList = UBOUND(aData) - LBOUND(aData) + 1
		REDIM aList(UBOUND(aData))
						
		DIM	i
		DIM	j
		DIM	n
		DIM	sLine
		DIM	aLine
		DIM	sSuffix
		j = 0
		i = 0
		DO WHILE i <= UBOUND(aData)
			sLine = aData(i)
			aLine = SPLIT( sLine, vbTAB )
			n = InStrRev( aLine(1), ".", -1, vbTextCompare )
			IF 0 < n THEN
				sSuffix = LCASE(MID( aLine(1), n ))
				SELECT CASE sSuffix
				CASE ".gif", ".jpg", ".jpeg" ', ".png"
					dTemp = CDATE(aLine(5))
					IF gPictureDate < dTemp THEN gPictureDate = dTemp
					aList(j) = "picture.asp?file=" & REPLACE(aLine(2),"\","\\")
					j = j + 1
				END SELECT
			END IF
			i = i + 1
		LOOP
		REDIM PRESERVE aList(j-1)
		nList = j
	ELSE
		REDIM aList(0)
		nList = 0
	END IF

END SUB



SUB getGalleryFilenames()

	getPictureFilenames g_aGalleryFilenames, g_nGalleryIndexCount, "Gallery/Scenery", g_bGalleryRandomize
	
END SUB

SUB getHomeFilenames()

	getPictureFilenames g_aGallery, g_nGallery, "Gallery/Scenery", TRUE
	
END SUB


DIM	g_aGallery()
DIM	g_nGallery
g_nGallery = 0


getHomeFilenames




%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" >
<title>Contact - Kodiak Photography</title>
<meta name="keywords" content="photography, photographer, huntsville, alabama, wedding, portrat, scenery">
<meta name="navigate" content="tab" >
<meta name="sortname" content="zzcontacts" >
<meta name="navtitle" content="Contact">
<!--#include file="scripts\favicon.asp"-->
<script type="text/javascript" src="scripts/slideshow.js"></script>
<script type="text/javascript" language="javascript" src="scripts/pobox.asp"></script>
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<style type="text/css">

#content
{
	text-align: center;
	font-family: Arial, Verdana, Helvetica, sans-serif;
	font-size: larger;
	position: relative;
	width: 100%;
}

#addressblock
{
	text-align: left;
	margin-top: 10px;
	margin-bottom: 10px;
	font-weight: bold;
	margin-right: auto;
	margin-left: auto;
}


#content .dialed a,
#content .dialed a:link,
#content .dialed a:visited,
#content .dialed a:active
{
	font-size: x-large;
}


#slideGallery
{
	position: relative;
	height: 400px;
	width: 400px;
	filter: blendTrans(Duration=1.0);
	overflow: hidden;
	text-align: center;
}








</style>
<script type="text/javascript">
<!--


<%

SUB makeJSImageList( sName, aList, nList )

	DIM	nMaxList
	nMaxList = nList
	IF 49 < nMaxList THEN nMaxList = 50

	Response.Write "var g_n" & sName & " = " & nMaxList & ";" & vbCRLF
	Response.Write "var g_a" & sName & " = new Array(" & nMaxList & ");" & vbCRLF
	'Response.Write "var g_a" & sName & "Cache = new Array(" & nMaxList & ");" & vbCRLF

	DIM xFile
	DIM	i
	i = 0
	FOR EACH xFile IN aList
		Response.Write "g_a" & sName & "[" & i & "] = """ & xFile & """;" & vbCRLF
		i = i + 1
		IF nMaxList < i THEN EXIT FOR
	NEXT 'xFile

END SUB

makeJSImageList "Gallery", g_aGallery, g_nGallery


%>











var g_h = 0;
var g_w = 0;

function CSlideViewFactory()
{
}

CSlideViewFactory.prototype.makeSlide
= function( s )
{
	var sPrefix = s.substr(0,4);
	if ( "htm:" == sPrefix )
		return new CSlideHTML( s.substr(4) );
	else if ( "img:" == sPrefix )
		return new CSlidePicture( s.substr(4) + "&h=" + g_h + "&w=" + g_w );
	else
		return new CSlidePicture( s + "&h=" + g_h + "&w=" + g_w );
}


var g_oGallerySlideFactory;
var g_oGalleryCarousel;

var g_oGalleryProjector;

var g_oGallerySlideScreen;







function setupResize()
{
	if (window.addEventListener)
		window.addEventListener("resize", pageReload, false)
	else if (window.attachEvent)
		window.attachEvent("onresize", pageReload)
}





function start_slider()
{
	var myWidth = 0, myHeight = 0;

	var	h;
	var	w;
	var oH = document.getElementById("slideGallery");

	h = getDivHeight( oH );
	w = getDivWidth( oH );
	
	
	g_h = h;
	g_w = w;
	
	//return;
	g_oGallerySlideFactory = new CSlideViewFactory();
	g_oGalleryCarousel = new CSlideCarousel();
	g_oGalleryCarousel.setSlideFactory( g_oGallerySlideFactory );
	g_oGalleryCarousel.load( g_aGallery );

	g_oGalleryProjector = new CSlideProjectorShowSingle();
	g_oGalleryProjector.setCarousel( g_oGalleryCarousel );
	g_oGalleryProjector.setPicturePause( <%=g_nPictureDelay%> );

	
	g_oGallerySlideScreen = new CSlideScreenTransition( "slideGallery" );
	g_oGallerySlideScreen.initialize();
	//g_oGallerySlideScreen.setHeight( h );
	//g_oGallerySlideScreen.setWidth( w );
	g_oGalleryProjector.setScreen( g_oGallerySlideScreen );

	g_oGalleryCarousel.setCookerTimer( 1.0 );

	setTimeout( "g_oGalleryCarousel.precookCards()", 0.01*1000 );
	setTimeout( "g_oGalleryProjector.run()", 0.1*1000 );


	//setTimeout( "setupResize()", 0.5*1000 );

}



if (window.addEventListener)
	window.addEventListener("load", start_slider, false)
else if (window.attachEvent)
	window.attachEvent("onload", start_slider)
else
	setTimeout( "start_slider()", 2.0*1000 );






//-->
</script>
</head>

<body>
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="scripts\include_navigation.asp"-->
<%
'outputSubtitleBlock
%>

<div id="content">

<table id="addressblock" border="0">
<tr>
<td align="center" style="font-weight: normal; color: #808080">
<div id="slideGallery">
<img alt="Shop and Save With Us" src="images/ContactUs.jpg">
</div></td>
<td valign="middle">
<p>
122 Rachel Drive<br>
Huntsville, AL 35806<br>
<span class="dial">25 6*7 22*91 28</span><br>
<span class="pobox">in fo*Ko di ak Ph o to gr ap hy:c om</span>
</p>
</td>
</tr>
</table>
<%



%>
<div style="position: relative; width: 30em; text-align: center; margin-right: auto; margin-left: auto;">
<div>Supporting your photography needs in North Alabama including:</div>
<div>
Huntsville, Madison, Decatur, Athens, Gurley, Harvest, Toney, Scottsboro, 
Guntersville, Arab, Somerville, Hartselle, Hazel Green,
</div>
<div>and many others...</div>
</div>

</div>

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