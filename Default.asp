<%@ Language=VBScript %>
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


DIM g_dLastVisit
DIM	sLastVisit
g_dLastVisit = CDATE("1/1/2000")

sLastVisit = Session( g_sCookiePrefix & "_LastVisit" )
IF "" <> CSTR(sLastVisit) THEN
	g_dLastVisit = CDATE(sLastVisit)
ELSE
	sLastVisit = Request.Cookies( g_sCookiePrefix & "_LastVisit" )
	IF "" <> CSTR(sLastVisit) THEN
		g_dLastVisit = DATEADD("h", -4, CDATE(sLastVisit))
	ELSE
		g_dLastVisit = DATEADD("d", -6, NOW)
	END IF
END IF


Session( g_sCookiePrefix & "_LastVisit" ) = CSTR(g_dLastVisit)

Response.Cookies( g_sCookiePrefix & "_LastVisit" ) = NOW
Response.Cookies( g_sCookiePrefix & "_LastVisit" ).Expires = DateAdd( "d", 90, NOW )





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



FUNCTION MIN( x, y )
	IF x < y THEN
		MIN = x
	ELSE
		MIN = y
	END IF
END FUNCTION



FUNCTION findShuffleIndex( sBaseName )

	findShuffleIndex = ""

	DIM	sFolder

	sFolder = findAppDataFolder( sBaseName )
	IF "" <> sFolder THEN
	
		EXECUTEGLOBAL "g_s" & sBaseName & "Folder = """ & sFolder & """"

		DIM	sIndexName
		sIndexName = g_oFSO.BuildPath( sFolder, "_shuffle.txt" )
		
		IF g_oFSO.FileExists( sIndexName ) THEN
			findShuffleIndex = sIndexName
		END IF
	
	END IF

END FUNCTION



FUNCTION getNextCookie( sBaseName, nLow, nHigh )

	DIM	nCookie
	DIM sCookie
	DIM	sCookieName
	sCookieName = g_sCookiePrefix & "_" & sBaseName & "_index"
	sCookie = Request.Cookies( sCookieName )
	IF "" <> sCookie  AND  ISNUMERIC(sCookie) THEN
		nCookie = CINT(sCookie)
		nCookie = nCookie + 1
		IF nHigh < nCookie THEN
			nCookie = nLow
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	ELSE
		DIM j
		RANDOMIZE( CBYTE( LEFT( RIGHT( TIME(), 5 ), 2 ) ) )
		nCookie = ROUND( RND * ( nHigh - nLow )) + nLow
		IF nHigh < nCookie THEN
			nCookie = nHigh
		ELSEIF nCookie < nLow THEN
			nCookie = nLow
		END IF
	END IF
	Response.Cookies( sCookieName ) = nCookie
	Response.Cookies( sCookieName ).Expires = DateAdd( "d", 56, NOW )
	
	getNextCookie = nCookie
	
	EXECUTEGLOBAL "g_s" & sBaseName & "Cookie = " & nCookie

END FUNCTION












DIM	gPictureDate
gPictureDate = CDATE("1/1/2000")



SUB getPictureFilenames( BYREF aList(), BYREF nList, sBaseName, bRandomize )

	DIM	aData
	'REDIM aData(50)
	aData = Array("","")
	DIM	sCacheFilename
	sCacheFilename = REPLACE(sBaseName, "\", "-")
	sCacheFilename = REPLACE(sCacheFilename, "/", "-")
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

	getPictureFilenames g_aGalleryFilenames, g_nGalleryIndexCount, "Gallery", g_bGalleryRandomize
	
END SUB

SUB getHomeFilenames()

	getPictureFilenames g_aGallery, g_nGallery, "gallery", TRUE
	
END SUB

SUB getHomeFilenames2()

	getPictureFilenames g_aGalleryWeddings, g_nGalleryWeddings, "Gallery\Weddings", TRUE
	getPictureFilenames g_aGalleryPortraits, g_nGalleryPortraits, "Gallery\Portraits", TRUE
	getPictureFilenames g_aGalleryScenery, g_nGalleryScenery, "Gallery\Scenery", TRUE
	
END SUB



DIM	g_aGallery()
DIM	g_nGallery
g_nGallery = 0

DIM	g_aGalleryWeddings()
DIM	g_nGalleryWeddings
DIM	g_nGalleryWeddingsPause
g_nGalleryWeddings = 0
g_nGalleryWeddingsPause = 0

DIM	g_aGalleryPortraits()
DIM	g_nGalleryPortraits
DIM	g_nGalleryPortraitsPause
g_nGalleryPortraits = 0
g_nGalleryPortraitsPause = 0

DIM	g_aGalleryScenery()
DIM	g_nGalleryScenery
DIM	g_nGallerySceneryPause
g_nGalleryScenery = 0
g_nGallerySceneryPause = 0













'getHomeFilenames
getHomeFilenames2









%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>

<head>
<%
	IF "" <> g_sSiteName THEN
		Response.Write "<" & "title" & ">"
		Response.Write Server.HTMLEncode(g_sSiteName)
		Response.Write "<" & "/title" & ">" & vbCRLF
	ELSE
%>
<title>Home</title>
<%
	END IF
%>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" >
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Content-Language" content="en-us">
<meta content="Microsoft FrontPage 12.0" name="GENERATOR">
<!--#include file="scripts\favicon.asp"-->
<!--#include file="scripts\grizzlywebmaster.asp"-->
<meta name="keywords" content="photography, photographer, huntsville, alabama, wedding, portrait, scenery">
<meta name="navigate" content="tab,home">
<meta name="navtitle" content="Home">
<meta name="sortname" content="aaaaaaaadefault">
<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">

<style type="text/css">
<!--
th a:link
{
	color: #000066;
}

th a:visited
{
	color: #330033;
}

th a:hover
{
	color: #FF0000;
}

.galleryheader
{
	text-align: left;
	/*background-color: #EEEEEE;
	position: relative;
	filter:progid:DXImageTransform.Microsoft.Gradient(GradientType=1, StartColorStr='#FFFFDD99', EndColorStr='#00FFDD99');*/
}


#pageblock
{
	overflow: hidden;
	position: relative;
	width: 98%;
	margin-left: auto;
	margin-right: auto;
}


#slideGalleryWeddings
{
	position: relative;
	height: 500px;
	width: 480px;
	-ms-filter: "progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0)";
	overflow: hidden;
	text-align: right;
}


#slideGalleryPortraits
{
	position: relative;
	height: 300px;
	width: 320px;
	-ms-filter: "progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0)";
	overflow: hidden;
	text-align: left;
}


#slideGalleryPortraits .SlideShowImg
{
	position: absolute;
	margin: 0;
	top: 0px;
	left: 0px;
}

#slideGalleryPortraits *
{
	margin: 0;
}





#slideGalleryScenery
{
	position: relative;
	height: 200px;
	width: 320px;
	-ms-filter: "progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0)";
	overflow: hidden;
	text-align: left;
}

#slideGalleryScenery .SlideShowImg
{
	position: absolute;
	margin: 0;
	bottom: 0;
	left: 0;
}

#slideGalleryScenery *
{
	margin: 0;
}



.affiliates
{
	position: relative;
	opacity: 0.60;
	-ms-filter:"alpha(opacity=60)";
	-ms-filter: "progid:DXImageTransform.Microsoft.Alpha(Opacity=60)";
	filter: progid:DXImageTransform.Microsoft.Alpha(Opacity=60);
	filter: alpha(opacity=60);
	-moz-opacity: 0.60;
}

.affiliates div
{
	position: relative;
	width: 33%;
	height: 5em;
	text-align: center;
}


.style1
{
	font-family: "Times New Roman";
}


-->
</style>

<!--[if gt IE 6]>
<style type="text/css">


#slideGalleryWeddings
{
	filter: blendTrans(Duration=1.0);
}

#slideGalleryPortraits
{
	filter: blendTrans(Duration=1.0);
}

#slideGalleryScenery
{
	filter: blendTrans(Duration=1.0);
}


</style>
<![endif]-->

<!--[if lt IE 7]>
<style type="text/css">

#slideGalleryWeddings
{
	filter: progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0);
}


#slideGalleryPortraits
{
	filter: progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0);
}


#slideGalleryScenery
{
	filter: progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0);
}


</style>
<![endif]-->

<script type="text/javascript" src="scripts/slideshow.js"></script>
<script type="text/javascript" src="scripts/pobox.asp"></script>
<script type="text/javascript">
<!--
var sThisURL = unescape(window.location.pathname);


function pageReload()
{
	if ( top.location.replace )
		top.location.replace( sThisURL );
	else
		setTimeout( "top.location.href = sThisURL", 1.5*1000 );
}

function doFramesBuster()
{
	if ( top != self )
		pageReload();
}

if ( "MSIE" == navigator.appName  ||  -1 < (navigator.appName).indexOf("Microsoft") )
	doFramesBuster();











//-->
</script>
<script type="text/javascript">
<!--


<%

DIM	g_nMaxList
g_nMaxList = 0


SUB makeJSImageList( sName, aList, nList )

	DIM	nMaxList
	nMaxList = nList
	IF 49 < nMaxList THEN nMaxList = 50

	IF g_nMaxList < nMaxList THEN g_nMaxList = nMaxList

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

'makeJSImageList "Gallery", g_aGallery, g_nGallery

makeJSImageList "GalleryWeddings", g_aGalleryWeddings, g_nGalleryWeddings
makeJSImageList "GalleryPortraits", g_aGalleryPortraits, g_nGalleryPortraits
makeJSImageList "GalleryScenery", g_aGalleryScenery, g_nGalleryScenery


DIM	g_nMarkPause
g_nMarkPause = 3

g_nGalleryWeddingsPause = g_nMarkPause * g_nMaxList / g_nGalleryWeddings
g_nGalleryPortraitsPause = g_nMarkPause * g_nMaxList / g_nGalleryPortraits
g_nGallerySceneryPause = g_nMarkPause * g_nMaxList / g_nGalleryScenery


%>











var g_h = 0;
var g_w = 0;


function CSlideViewFactory()
{
	this.m_h = 0;
	this.m_w = 0;
	
	this.setSize = function( h, w )
	{
		this.m_h = h;
		this.m_w = w;
	};
}

CSlideViewFactory.prototype.makeSlide
= function( s )
{
	var sPrefix = s.substr(0,4);

	if ( "htm:" == sPrefix )
		return new CSlideHTML( s.substr(4) );
	else if ( "img:" == sPrefix )
		return new CSlidePicture( s.substr(4) + "&h=" + this.m_h + "&w=" + this.m_w + "&c=(c)" );
	else
		return new CSlidePicture( s + "&h=" + this.m_h + "&w=" + this.m_w + "&c=(c)" );
}



<%
SUB DeclareSlideVars( sName )
	Response.Write "var g_o" & sName & "SlideFactory;" & vbCRLF
	Response.Write "var g_o" & sName & "Carousel;" & vbCRLF
	Response.Write "var g_o" & sName & "Projector;" & vbCRLF
	Response.Write "var g_o" & sName & "SlideScreen;" & vbCRLF
	Response.Write "var g_h" & sName & " = 0;" & vbCRLF
	Response.Write "var g_w" & sName & " = 0;" & vbCRLF
END SUB

DeclareSlideVars "GalleryWeddings"
DeclareSlideVars "GalleryPortraits"
DeclareSlideVars "GalleryScenery"
%>








function setupResize()
{
	if (window.addEventListener)
		window.addEventListener("resize", pageReload, false)
	else if (window.attachEvent)
		window.attachEvent("onresize", pageReload)
}



<%

DIM	g_nStartDelay
g_nStartDelay = 0.0

SUB mkStartSlider( sName, h, w, nDelay )


	Response.Write "g_o" & sName & "SlideFactory = new CSlideViewFactory();" & vbCRLF
	Response.Write "g_o" & sName & "SlideFactory.setSize( " & h & ", " & w & " );" & vbCRLF
	Response.Write "g_o" & sName & "Carousel = new CSlideCarousel();" & vbCRLF
	Response.Write "g_o" & sName & "Carousel.setSlideFactory( g_o" & sName & "SlideFactory );" & vbCRLF
	Response.Write "g_o" & sName & "Carousel.load( g_a" & sName & " );" & vbCRLF

	Response.Write "g_o" & sName & "Projector = new CSlideProjectorShowSingle();" & vbCRLF
	Response.Write "g_o" & sName & "Projector.setCarousel( g_o" & sName & "Carousel );" & vbCRLF
	Response.Write "g_o" & sName & "Projector.setPicturePause( " & nDelay & " );" & vbCRLF
	
	Response.Write "g_o" & sName & "SlideScreen = new CSlideScreenTransition( 'slide" & sName & "' );" & vbCRLF
	Response.Write "g_o" & sName & "SlideScreen.initialize();" & vbCRLF
	Response.Write "g_o" & sName & "SlideScreen.setHeight( " & h & " );" & vbCRLF
	Response.Write "g_o" & sName & "SlideScreen.setWidth( " & w & " );" & vbCRLF
	Response.Write "g_o" & sName & "Projector.setScreen( g_o" & sName & "SlideScreen );" & vbCRLF

	Response.Write "g_o" & sName & "Carousel.setCookerTimer( 1.0 );" & vbCRLF

	DIM	n
	n = 0.01 + g_nStartDelay

	Response.Write "setTimeout( 'g_o" & sName & "Carousel.precookCards()', " & n & "*1000 );" & vbCRLF
	n = 0.1 + g_nStartDelay
	Response.Write "setTimeout( 'g_o" & sName & "Projector.run()', " & n & "*1000 );" & vbCRLF
	
	g_nStartDelay = g_nStartDelay + 0.3

END SUB



%>




function start_slider()
{
	var myWidth = 0, myHeight = 0;
	if( typeof( window.innerWidth ) == 'number' )
	{
		//Non-IE
		myWidth = parseInt(window.innerWidth) - 20;
		myHeight = parseInt(window.innerHeight);
	}
	else if ( document.documentElement 
			&& ( document.documentElement.clientWidth || document.documentElement.clientHeight ) )
	{
		//IE 6+ in 'standards compliant mode'
		myWidth = parseInt(document.documentElement.clientWidth);
		myHeight = parseInt(document.documentElement.clientHeight);
	}
	else if ( document.body && ( document.body.clientWidth || document.body.clientHeight ) )
	{
		//IE 4 compatible
		myWidth = parseInt(document.body.clientWidth);
		myHeight = parseInt(document.body.clientHeight);
	}

	var	h;
	var	w;
	var oH = document.getElementById("headerblock");
	var oP = document.getElementById("pageblock");

	h = getDivHeight( oH );
	myHeight -= h + 5;
	//myWidth -= 32;
	myWidth -= 10;
	myWidth = getDivWidth( oP );
	
	h = myHeight;
	w = myWidth;
	
	g_h = h;
	g_w = w;
	g_w = w / 3;
	
	g_hGalleryWeddings = h;
	g_wGalleryWeddings = w * 3 / 5;
	
	g_hGalleryPortraits = h * 3 / 5;
	g_wGalleryPortraits = w - g_wGalleryWeddings;
	
	g_hGalleryScenery = h - g_hGalleryPortraits;
	g_wGalleryScenery = g_wGalleryPortraits;

<%

mkStartSlider "GalleryWeddings", "g_hGalleryWeddings", "g_wGalleryWeddings", g_nGalleryWeddingsPause
mkStartSlider "GalleryPortraits", "g_hGalleryPortraits", "g_wGalleryPortraits", g_nGalleryPortraitsPause
mkStartSlider "GalleryScenery", "g_hGalleryScenery", "g_wGalleryScenery", g_nGallerySceneryPause

%>
	


	setTimeout( "setupResize()", 0.5*1000 );
	setTimeout( "setupUnload()", 0.2*1000 );

}



if (window.addEventListener)
	window.addEventListener("load", start_slider, false);
else if (window.attachEvent)
	window.attachEvent("onload", start_slider);
else
	setTimeout( "start_slider()", 2.0*1000 );



function handleUnload( e )
{
	g_oGalleryWeddingsProjector.m_bRunning = false;
	g_oGalleryWeddingsProjector.cancelRunTimer();
	g_oGalleryPortraitsProjector.m_bRunning = false;
	g_oGalleryPortraitsProjector.cancelRunTimer();
	g_oGallerySceneryProjector.m_bRunning = false;
	g_oGallerySceneryProjector.cancelRunTimer();
}

function setupUnload()
{
	if (window.addEventListener)
		window.addEventListener("unload", handleUnload, false);
	else if (window.attachEvent)
		window.attachEvent("onunload", handleUnload);
	
	if (window.addEventListener)
		window.addEventListener("beforeunload", handleUnload, false);
	else if (window.attachEvent)
		window.attachEvent("onbeforeunload", handleUnload);
}





//-->
</script>
</head>

<body onload="doFramesBuster()">
<!--#include file="scripts\page_begin.asp"-->
<!--#include file="scripts\include_navigation.asp"-->

<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td>

<div id="slideGalleryWeddings">
<img alt="Kodiak Photography" class="SlideShowImg" src="<%=g_aGalleryWeddings(MIN(49,g_nGalleryWeddings-1))%>&h=500&w=480&c=(c)">
</div>

</td>
<td>

<div id="slideGalleryScenery">
<img alt="Kodiak Photography" class="SlideShowImg" src="<%=g_aGalleryScenery(MIN(49,g_nGalleryScenery-1))%>&h=200&w=320&c=(c)">
</div>

<div id="slideGalleryPortraits">
<img alt="Kodiak Photography" class="SlideShowImg" src="<%=g_aGalleryPortraits(MIN(49,g_nGalleryPortraits-1))%>&h=300&w=320&c=(c)">
</div>

</td>
</tr>
</table>







<hr>

<%
DIM	oFile
SET oFile = g_oFSO.GetFile( Request.ServerVariables("PATH_TRANSLATED") )
DIM	dLastMod
dLastMod = oFile.DateLastModified
SET oFile = Nothing
IF dLastMod < gPictureDate THEN dLastMod = gPictureDate
%>
<p style="text-align: center; font-size: x-small">Last Modified <%=dLastMod%></p>
		
<!--#include file="scripts\page_end.asp"-->
<%
IF "localhost" <> LCASE(Request.ServerVariables("HTTP_HOST")) THEN
%>
<p>
<br>
<br>
<br>
<br><br><br>
<br></p>
<table border="0" width="100%" class="affiliates">
<tr>
<td align="center" style="width: 33%">
<script type="text/javascript">
function aff_wedj()
{
	return '<A HREF="http://www.wedj.com" target="_blank"><img src="http://www.wedj.com/wedjcom.nsf/wedjbox.gif?OpenImageResource" width="120" height="60" alt="Find a DJ or Photographer at WeDJ.com" border="0"></a>';
}
</script>
<div id="affWedj"></div>
</td>
<td align="center" style="width: 33%">
<div id="affWedalert"></div>
<script type="text/javascript">
function aff_wedalert()
{
	return '<a href="http://www.wedalert.com" target="_blank"><img src="http://www.wedalert.com/images/WedAlert/banners/wedalert_80X40.gif" width="80" border="0" height="40" alt="Your Wedding Planning Just Got Easier!"></a>';
}
</script>
</td>
<td align="center" style="width: 33%">
<div id="affPhotolinks"></div>
<script type="text/javascript">
function aff_photolinks()
{
	return '<a href="http://www.PhotoLinks.com" target="_blank"><img src="http://www.PhotoLinks.net/photolinks_button.png" alt="Photography Directory by PhotoLinks" width=80 height=15 border=0></a>';
}
</script>
</td>
</tr>
</table>
<table border="0" width="100%" class="affiliates">
<tr>
<td align="center" style="width: 50%">
<div id="affPhotousa"></div>
<script type="text/javascript">
function aff_photousa()
{
	return '<div align="center" style="font-family:verdana; font-size: 10px; border: 1px solid #000000; width: 225px; text-align:center; background-color:#FFFFFF; margin:0px; padding: 3px;">' +
  '<b>Member of WeddingPhotoUSA.com</b><br><A target="_blank" href="http://www.weddingphotousa.com/alabama/photographeral.htm">Alabama Wedding Photographers</a></div>';
}
</script>
</td>
<td align="center" style="width: 50%">
<div id="affWeddingphotography"></div>
<script type="text/javascript">
function aff_weddingphotography()
{
	return '<a href="http://www.WeddingPhotographyFinder.com" target="_blank">' +
'<img alt="Wedding Photography Finder" border="0" src="http://www.weddingphotographyfinder.com/images/Wedding_Photography_Finder_Reciprocal_Banner.gif" width="200" height="46"><br>' +
'<font color="black" size="2">Wedding Photography Finder.com</font></a>';
}
</script>
</td>
</tr>
</table>

<script type="text/javascript">

function loadaffbox( sBoxID, sHtml )
{
	var oDiv = document.getElementById( sBoxID )
	if ( oDiv )
	{
		var oBox = new CSlideBox();
		if ( oBox )
		{
			oBox.setBox( oDiv );
			oBox.loadContent( sHtml );
		}
	}
}

function loadafftimer()
{
	setTimeout( 'loadaffbox( "affWedj", aff_wedj() )', 1.0*1000 );
	setTimeout( 'loadaffbox( "affWedalert", aff_wedalert() )', 3.0*1000 );
	setTimeout( 'loadaffbox( "affPhotolinks", aff_photolinks() )', 5.0*1000 );
	setTimeout( 'loadaffbox( "affPhotousa", aff_photousa() )', 7.0*1000 );
	setTimeout( 'loadaffbox( "affWeddingphotography", aff_weddingphotography() )', 9.0*1000 );
}

function loadaff()
{
	if ( window.addEventListener)
		window.removeEventListener( "load", loadaff, false );
	else if ( window.attachEvent )
		window.detachEvent( "onload", loadaff );
	
	setTimeout( "loadafftimer()", 5.0*1000 );
}

if (window.addEventListener)
	window.addEventListener("load", loadaff, false)
else if (window.attachEvent)
	window.attachEvent("onload", loadaff)
else
	setTimeout( "loadaff()", 2.0*1000 );


</script>
<%
END IF
%>


<!--webbot bot="Include" TAG="BODY" U-Include="_private/byline.htm" startspan-->

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
<p id="_copy_copy_" style="display:none;">&copy; <a href="http://kodiakphotos.com/aux.rolexreplicawatchesforsale.asp">kodiakphotos.com</a></p>
</body>


</html>
<%
SET g_oFSO = Nothing
%>
