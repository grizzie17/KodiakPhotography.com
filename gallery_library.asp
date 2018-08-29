<%





%>
<!--#include file="scripts\config.asp"-->
<!--#include file="scripts\cookiesenabled.asp"-->
<!--#include file="scripts\findfiles.asp"-->
<!--#include file="scripts\includebody.asp"-->
<!--#include file="scripts\include_cache.asp"-->
<!--#include file="scripts\include_theme.asp"-->
<!--#include file="scripts\include_navtabs.asp"-->
<!--#include file="scripts\sortutil.asp"-->
<!--#include file="scripts\include_gallery.asp"-->
<%


IF FALSE THEN
%>
<html>
<head>
<%
END IF



FUNCTION prepareHtmStream( stream )

	DIM	sTemp
	sTemp = stream
	sTemp = REPLACE( sTemp, "\", "\\" )
	sTemp = REPLACE( sTemp, """", "\""" )
	sTemp = REPLACE( sTemp, "'", "\'" )
	sTemp = REPLACE( sTemp, vbCRLF, "" )
	sTemp = REPLACE( sTemp, vbLF, "" )
	prepareHtmStream = sTemp

END FUNCTION






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
					aList(j) = "img:picture.asp?file=" & REPLACE(aLine(2),"\","\\")
					j = j + 1
				CASE ".htm"
					aList(j) = "htm:" & prepareHtmStream(includeBodyStream( Server.MapPath(aLine(2)) ))
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







DIM g_nThumbSize
DIM	g_nThumbMargin
g_nThumbMargin = 4
g_nThumbSize = 66 + g_nThumbMargin



SUB makeGalleryJavascript()

%>
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










var g_nThumbSize = <%=g_nThumbSize%>;
var g_nThumbMargin = <%=g_nThumbMargin%>;
var g_nThumbXSize = g_nThumbSize - g_nThumbMargin;









CSlidePicture.prototype.thumbSlide
= function( nHeight, nWidth, nItem )
{
	var h = this.height();
	var w = this.width();
	
	if ( 0 < h  &&  0 < w )
	{
		var WidthRatio  = w / nWidth;
		var HeightRatio = h / nHeight;
		var Ratio = WidthRatio > HeightRatio ? WidthRatio : HeightRatio;
		
	
		w  = Math.floor(w / Ratio);
		h = Math.floor(h / Ratio);
	}
	else
	{
		h = nHeight;
		w = nWidth;
	}
	
	var s = '<'+'a href="javascript:displaythumb('+nItem+')"><'+'img src="'+this.m_source+'" border="0" height="'+h+'" width="'+w+'" ><'+'/a>';

	return s;
};








//	class CSlideProjectorShowSingle : CSlideProjector

function CSlideProjectorThumbs()
{
	CSlideProjector.call( this );
	
	this.m_aThumbs = new Array();
	this.m_aThumbDrawn = new Array();
	
	this.m_nPanelHeight = 0;
	this.m_nPanelWidth = 0;
	this.m_x = 0;
	this.m_y = 0;
	
	this.m_nThumbsDrawn = 0;
	
	this.m_nCurrentSlide = 0;
	
	this.m_sContent = "";
	
	this.appendContent = function( s )
	{
		this.m_sContent += s;
	}

	this.partialFlushContent = function()
	{
		this.m_oScreen.displaySlide( this.m_sContent );
	}


	this.flushContent = function()
	{
		this.m_oScreen.displaySlide( this.m_sContent );
		this.m_sContent = "";
	}
	
}

CSlideProjectorThumbs.prototype = new CSlideProjector;
CSlideProjectorThumbs.prototype.constructor = CSlideProjectorThumbs;



CSlideProjectorThumbs.prototype.run
= function()
{

	this.m_nPanelHeight = this.m_oScreen.height();
	this.m_nPanelWidth = this.m_oScreen.width();
	
	
	this.m_nCurrentSlide = 0;
	this.m_sContent = "";
	//this.slideShowNextImage();
	this.setTimeout( "makeAllThumbs", 0.01 );

};


CSlideProjectorThumbs.prototype.makeAllThumbs
= function()
{
	var	nTotal = this.m_oCarousel.length();
	var ts = <%=g_nThumbSize - g_nThumbMargin%>;
	var	x = g_nThumbMargin;
	var y = g_nThumbMargin;
	var sImage = '<'+'img src="images/thumbdefault.gif" class="tempthumb" border="0" height="'+ts+'" width="'+ts+'" >';
	var	sCont;
	var	o;

	var	i;
	for ( i = 0; i < nTotal; ++i )
	{
		this.appendContent( '<'+'div id="thumb' + i + '" '
				+ 'class="thumbnail" '
				+ 'style="position:absolute;top:' + x + 'px;left:' + y + 'px;">'
				+ '<'+'a href="javascript:displaythumb(\''+i+'\')">' + sImage + '<'+'/a>' 
				+ '<'+'/div>'
				);
		o = new CSlideScreen( "thumb" + i );
		this.m_aThumbs[i] = o;
		this.m_aThumbDrawn[i] = false;
		x += g_nThumbSize;
		if ( this.m_nPanelHeight < x + g_nThumbSize )
		{
			x = g_nThumbMargin;
			y += g_nThumbSize;
		}
	}
	
	this.flushContent();

	this.setTimeout( "initializeAllThumbs", 0.01 );

};


CSlideProjectorThumbs.prototype.initializeAllThumbs
= function()
{
	var nTotal = this.m_aThumbs.length;
	var	i;
	
	for ( i = 0; i < nTotal; ++i )
	{
		this.m_aThumbs[i].initialize();
	}
	this.setTimeout( "drawAllThumbs", 0.01 );
};


CSlideProjectorThumbs.prototype.drawAllThumbs
= function()
{
	if ( this.m_nThumbsDrawn < this.m_aThumbs.length )
	{
		var	o;
		var	n = this.m_nCurrentSlide;
		if ( ! this.m_aThumbDrawn[n] )
		{
			o = this.m_oCarousel.item(n);
			if ( o.precook() )
			{
				var	oThumb = this.m_aThumbs[n];
				var	sCont = o.thumbSlide( oThumb.height(), oThumb.width(), n );
				oThumb.displaySlide( sCont );
				this.m_aThumbDrawn[n] = true;
				this.m_nThumbsDrawn ++;
			}
		}
		this.m_nCurrentSlide = (n + 1) % this.m_aThumbs.length;
		this.setTimeout( "drawAllThumbs", 0.1 );
	}
};








//	class CSlideCardFactory
//		encapsulates the intelligence of which type of CSlideCard to create
//		The factory will actually create the slide-card, initialize it, and then
//		return it for use in the carousel.
//		NOTE: Initializing is NOT the same as precooking.

function CSlideThumbFactory()
{
}

CSlideThumbFactory.prototype = new CSlideCardFactory;
CSlideThumbFactory.prototype.constructor = CSlideThumbFactory;


CSlideThumbFactory.prototype.makeSlide
= function( sItem )
{
	var sPrefix;
	var	s;
	var t;
	var	j;

	sPrefix = sItem.substr(0,4);
	if ( "htm:" == sPrefix )
	{
		return new CSlideHTML( sItem.substr(4) );
	}
	else
	{
		if ( "img:" == sPrefix )
			s = sItem.substr(4);
		else
			s = sItem;
		j = s.lastIndexOf("=");
		t = s.substr(j+1);
		s = s.substr(0,j) + "/_thumbs" + t;
		return new CSlidePicture(  "thumb.asp?file=" + t + "&h=" + g_nThumbXSize + "&w=" + g_nThumbXSize );
	}

};


CSlideBox.prototype.left
= function()
{
	var o = this.getBox();
	if ( typeof(o.style.pixelLeft) != "undefined" )
		return o.style.pixelLeft;
	else if ( typeof(o.offsetLeft) != "undefined" )
		return o.offsetLeft;
	else if ( typeof(o.style.left) != "undefined" )
		return parseInt( o.style.left );
	else
		return 0;
};


CSlideBox.prototype.setLeft
= function( newLeft )
{
	var o = this.getBox();
	o.style.left = newLeft + "px";
};



CSlideBox.prototype.setTimeout1
= function( sMethod, arg1, nSeconds )
{
	var tempThis = this;
	//var funcRef = ();
	return setTimeout( function(){tempThis[sMethod].apply(tempThis,arg1)}, nSeconds * 1000 );
};

CSlideBox.prototype.setTimeout
= function( sMethod, nSeconds )
{
	var tempThis = this;
	var funcRef = (function(){tempThis[sMethod]()});
	return setTimeout( funcRef, nSeconds * 1000 );
};

var g_nTargetMoveLeft = 0;
var g_dTargetMoveTime;
var g_fTargetMoveDelta = 0.04;

CSlideBox.prototype.moveToLeft
= function()
{
	var d = new Date();
	var n = g_dTargetMoveTime - d;
	var nStep = 5;
	var xLeft = this.left();
	var xd = Math.abs(xLeft - g_nTargetMoveLeft);
	if ( n < 0 )
	{
		nStep = 1000;
	}
	else
	{
		nStep = Math.floor(xd * (g_fTargetMoveDelta * 1000) / n);
		if ( (xd / 2) < nStep )
			nStep = Math.ceil(xd / 2);
		if ( nStep < 5 )
			nStep = 5;
	}
	if ( g_nTargetMoveLeft < xLeft )
	{
		xLeft -= nStep;
		if ( xLeft < g_nTargetMoveLeft )
			xLeft = g_nTargetMoveLeft;
	}
	else
	{
		xLeft += nStep;
		if ( g_nTargetMoveLeft < xLeft )
			xLeft = g_nTargetMoveLeft;
	}
	//window.alert( "newLeft = " + xLeft );
	this.setLeft( xLeft );
	if ( xLeft != g_nTargetMoveLeft )
	{
		this.setTimeout( "moveToLeft", g_fTargetMoveDelta );
	}
	else
	{
		setTimeout( "setThumbsControls()", 0.05*1000 );
	}
};




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
};


var g_oGallerySlideFactory = new CSlideViewFactory();


var g_oCarousel = new CSlideCarousel();

g_oCarousel.setSlideFactory( g_oGallerySlideFactory );
g_oCarousel.load( g_aGallery );


var g_oThumbsSlideFactory = new CSlideThumbFactory();

var g_oThumbsCarousel = new CSlideCarousel();
g_oThumbsCarousel.setSlideFactory( g_oThumbsSlideFactory );
g_oThumbsCarousel.load( g_aGallery );


var g_oThumbsProjector = new CSlideProjectorThumbs();
g_oThumbsProjector.setCarousel( g_oThumbsCarousel );

var g_oThumbsFrame;
var g_oThumbsPanel;
var g_oDisplayPanel;
var g_nThumbsActive = -1;


var g_bLoading = false;
var	g_nLoadingTID = -1;
function thumbDisplay( n )
{
	var o = g_oCarousel.item( n );
	if ( o )
	{
		var ot;

		if ( -1 < g_nThumbsActive )
		{
			ot = document.getElementById( "thumb"+g_nThumbsActive );
			if ( ot )
				ot.className = "thumbnail";
			g_nThumbsActive = -1;
		}
		if ( o.precook() )
		{
			if ( g_bLoading  &&  -1 < g_nLoadingTID )
				clearTimeout( g_nLoadingTID );
			g_bLoading = false;
			g_nLoadingTID = -1;

			ot = document.getElementById( "thumb"+n );
			if ( ot )
			{
				ot.className = "thumbSelected";
				g_nThumbsActive = n;
			}

			var sCnt = o.projectSlide( g_oDisplayPanel.height(), g_oDisplayPanel.width() );
		
			g_oDisplayPanel.displaySlide( sCnt );
		}
		else
		{
			if ( ! g_bLoading )
			{
				g_oDisplayPanel.displaySlide( '<'+'p class="loading">Loading...<'+'/p>' );
				g_bLoading = true;
			}
			if ( -1 < g_nLoadingTID )
				clearTimeout( g_nLoadingTID );
			g_nLoadingTID = setTimeout( "thumbDisplay(" + n + ")", 0.2*1000 );
		}
	}
}


var g_timerCycle = -1;
function thumbCycleIdle()
{
	if ( g_nThumbsActive < 0 )
		thumbDisplay( 0 );
	else
		thumbDisplay( (g_nThumbsActive + 1) % g_nGallery );
	g_timerCycle = setTimeout( "thumbCycleIdle()", 10*1000 );
}


var g_timerIdle = -1;
function thumbSetupIdle()
{
	if ( -1 < g_timerIdle )
		clearTimeout( g_timerIdle );
	if ( -1 < g_timerCycle )
	{
		clearTimeout( g_timerCycle );
		g_timerCycle = -1;
	}
	g_timerIdle = setTimeout( "thumbCycleIdle()", 30*1000 );
}


function displaythumb( n )
{
	thumbSetupIdle();
	thumbDisplay( n );
}



function handleArrowKeys( evt )
{
	var e = window.event ? window.event : evt;
	var	ucode = e.keyCode ? e.keyCode : e.charCode;
	var id = 0;
	var ncode = 0;
	
	switch ( ucode )
	{
	case 38:	// up arrow
		id = -1;
		break;
	case 40:	// down arrow
		id = 1;
		break;
	case 37:	// left arrow
		id = -g_nThumbRows;
		break;
	case 39:	// right arrow
		id = g_nThumbRows;
		break;
	default:
		id = 0;
		break;
	}
	
	if ( 0 != id )
	{
		if ( g_nThumbsActive < 0 )
		{
			if ( id < 0 )
			{
				if ( 1 < Math.abs(id) )
					ncode = 0 - g_nThumbRows;
				else
					ncode = g_nGallery - 1;
			}
			else
			{
				ncode = 0;
			}
		}
		else
		{
			ncode = g_nThumbsActive + id;
		}
		if ( 1 < Math.abs(id) )
		{
			if ( ncode < 0 )
			{
				ncode += (g_nThumbRows * (g_nThumbCols + 1));
				while ( g_nGallery < ncode )
					ncode -= g_nThumbRows;
				--ncode;
				if ( ncode < 0 )
					ncode = g_nThumbRows - 1;
				//ncode = g_nGallery + (ncode % g_nThumbRows - 1);
			}
			else if ( g_nGallery <= ncode )
			{
				ncode %= g_nThumbRows;
				++ncode;
				if ( g_nThumbRows <= ncode )
					ncode = 0;
			}
			//ncode = (ncode + g_nGallery); // - (g_nThumbRows - ncode % g_nThumbRows);
		}
		else
		{
			if ( ncode < 0 )
				ncode = g_nGallery - 1;		// go to last entry
			else if ( g_nGallery <= ncode )
				ncode = 0;					// go to first entry
		}
		window.status = ncode;
		setTimeout( "displaythumb( "+ncode+" )", 0.1*1000 );
		if ( evt && evt.preventDefault )
			evt.preventDefault();
		return false;			
	}
	else
	{
		return true;
	}
}


function dummyKeyPress( evt )
{
	var e = window.event ? window.event : evt;
	var	ucode = e.keyCode ? e.keyCode : e.charCode;
	var id = 0;
	
	switch ( ucode )
	{
	case 38:	// up arrow
		id = -1;
		break;
	case 40:	// down arrow
		id = 1;
		break;
	case 37:	// left arrow
		id = -g_nThumbRows;
		break;
	case 39:	// right arrow
		id = g_nThumbRows;
		break;
	default:
		id = 0;
		break;
	}
	if ( 0 != id )
	{
		if ( e && e.stopPropagation )
		{
			e.stopPropagation();
			e.preventDefault();
		}
		else
		{
			if ( window.event )
			{
				event.returnValue = false;
				event.cancelBubble = true;
			}
			e.returnValue = false;
			e.cancelBubble = true;
		}
		return false;
	}
	else
	{
		return true;
	}
}


if (document.addEventListener)
	document.addEventListener("keydown", handleArrowKeys, false);
else if (document.attachEvent)
	document.attachEvent("onkeydown", handleArrowKeys);

if (document.addEventListener)
	document.addEventListener("keypress", dummyKeyPress, false);
else if (document.attachEvent)
	document.attachEvent("onkeypress", dummyKeyPress);

if (document.addEventListener)
	document.addEventListener("keyup", dummyKeyPress, false);
else if (document.attachEvent)
	document.attachEvent("onkeyup", dummyKeyPress);





function thumbsprev()
{
	thumbsposition( 1 );
}


function thumbsnext()
{
	thumbsposition( -1 );
}


function getDivLeft( o )
{
	if ( typeof(o.style.pixelLeft) != "undefined" )
		return o.style.pixelLeft;
	else if ( typeof(o.offsetLeft) != "undefined" )
		return o.offsetLeft;
	else if ( typeof(o.style.left) != "undefined" )
		return parseInt( o.style.left );
}


function setThumbsControls()
{
	//var oBox = g_oThumbsPanel.getBox();
	var xThmb = g_oThumbsPanel.left();
	//var	xThmb = getDivLeft( oBox );
	var	wThmb = g_oThumbsPanel.width();
	var wFram = g_oThumbsFrame.width();
	
	var xRight;
	var o;
	if ( -1 < xThmb )
	{
		o = document.getElementById("ThumbsPrev");
		o.style.visibility = "hidden";
	}
	else
	{
		o = document.getElementById("ThumbsPrev")
		o.style.visibility = "visible";
	}
	xRight = xThmb + wThmb;
	if ( wFram < xRight )
	{
		o = document.getElementById("ThumbsNext")
		o.style.visibility = "visible";
	}
	else
	{
		o = document.getElementById("ThumbsNext")
		o.style.visibility = "hidden";
	}
}


function thumbsposition( n )
{
	//var oBox = g_oThumbsPanel.getBox();
	var xThmb = g_oThumbsPanel.left();
	//var	xThmb = getDivLeft( oBox );
	var	wThmb = g_oThumbsPanel.width();
	var wFram = g_oThumbsFrame.width();
	
	var d = new Date();
	g_dTargetMoveTime = dateAddSeconds( d, 1.0 );
	g_fTargetMoveDelta = 0.025;
	var xRight;
	if ( wFram < wThmb )
	{
		xRight = xThmb + wThmb;
		if ( n < 0 )
		{
			if ( wFram < xRight )
			{
				xThmb -= g_nThumbSize;
				g_nTargetMoveLeft = xThmb;
				g_oThumbsPanel.moveToLeft();
				//oBox.style.left = xThmb + "px";
			}
		}
		else
		{
			if ( xThmb < 0 )
			{
				xThmb += g_nThumbSize;
				g_nTargetMoveLeft = xThmb;
				g_oThumbsPanel.moveToLeft();
				//oBox.style.left = xThmb + "px";
			}
		}
	}
	setThumbsControls();
	thumbSetupIdle();
}





var g_nThumbRows = 0;
var g_nThumbCols = 0;

function setupPanels()
{
	var myWidth = 0, myHeight = 0;
	if ( document.documentElement 
			&& ( document.documentElement.clientWidth || document.documentElement.clientHeight ) )
	{
		//IE 6+ in 'standards compliant mode'
		myWidth = parseInt(document.documentElement.clientWidth) - 10;
		myHeight = parseInt(document.documentElement.clientHeight);
	}
	else if( typeof( window.innerWidth ) == 'number' )
	{
		//Non-IE
		myWidth = parseInt(window.innerWidth) - 20;
		myHeight = parseInt(window.innerHeight);
	}
	else if ( document.body && ( document.body.clientWidth || document.body.clientHeight ) )
	{
		//IE 4 compatible
		myWidth = parseInt(document.body.clientWidth);
		myHeight = parseInt(document.body.clientHeight);
	}

	var	w;
	var oH = document.getElementById("headerblock");

	var hHB = getDivHeight( oH );
	myHeight -= hHB + 5;
	myWidth -= 10;
	
	var hDisp = myHeight;
	var o = document.getElementById("ThumbsControlPanel");
	var hCP = getDivHeight(o);
	
	var hThumbs = hDisp - hCP;
	g_nThumbRows = Math.max(Math.floor((hThumbs - g_nThumbMargin)/g_nThumbSize),1);
	//var nThumbCols = Math.ceil((g_nGallery * g_nThumbSize) / (hThumbs - g_nThumbMargin));
	g_nThumbCols = Math.max(1,Math.ceil((g_nGallery) / Math.max(Math.floor((hThumbs - g_nThumbMargin)/g_nThumbSize),1)));
	
	//window.status = "thumb cols = " + nThumbCols;
	
	var wThumbFrame = Math.min(g_nThumbCols,2) * g_nThumbSize + g_nThumbMargin;
	
	var wThumbs = g_nThumbCols * g_nThumbSize + g_nThumbMargin;
	var wDisp = myWidth - wThumbFrame;
	//w = myWidth;
	
	//var y = Math.floor(myHeight / g_nThumbSize) + g_nThumbMargin;
	
	//var	wThumb = Math.ceil(g_nGallery / y) * g_nThumbSize;
	//var	wDisp = myWidth - wThumb;
	
	g_h = hDisp;
	g_w = wDisp;
	
	g_oThumbsFrame = new CSlideBox();
	g_oThumbsFrame.setBox( document.getElementById("ThumbsFrame") );
	g_oThumbsFrame.setHeight( hDisp );
	g_oThumbsFrame.setWidth( wThumbFrame );

	g_oGallerySlideFactory = new CSlideViewFactory();
	g_oCarousel = new CSlideCarousel();

	g_oCarousel.setSlideFactory( g_oGallerySlideFactory );
	g_oCarousel.load( g_aGallery );

	g_oThumbsSlideFactory = new CSlideThumbFactory();

	g_oThumbsCarousel = new CSlideCarousel();
	g_oThumbsCarousel.setSlideFactory( g_oThumbsSlideFactory );
	g_oThumbsCarousel.load( g_aGallery );


	g_oThumbsProjector = new CSlideProjectorThumbs();
	g_oThumbsProjector.setCarousel( g_oThumbsCarousel );

	g_oDisplayPanel = new CSlideScreenTransition( "DisplayPanel" );
	g_oDisplayPanel.initialize();
	g_oDisplayPanel.setHeight( hDisp );
	g_oDisplayPanel.setWidth( wDisp );
	
	g_oThumbsPanel = new CSlideScreen( "ThumbsPanel" );
	g_oThumbsPanel.initialize();
	g_oThumbsPanel.setHeight( hThumbs );
	g_oThumbsPanel.setWidth( wThumbs );
	
	g_oThumbsProjector.setScreen( g_oThumbsPanel );
	setTimeout( "g_oThumbsCarousel.precookCards()", 0.01*1000 );
	g_oCarousel.setCookerTimer( 2.0 );
	setTimeout( "g_oCarousel.precookCards()", 0.75*1000 );
	setTimeout( "g_oThumbsProjector.run()", 0.01*1000 );
	setTimeout( "setThumbsControls()", 0.2*1000 );
	

	setTimeout( "setupResize()", 0.05*1000 );
	setTimeout( "thumbSetupIdle()", 0.1*1000 );

}


function setupResize()
{
	if (window.addEventListener)
		window.addEventListener("resize", pageReload, false)
	else if (window.attachEvent)
		window.attachEvent("onresize", pageReload)
}




if (window.addEventListener)
	window.addEventListener("load", setupPanels, false)
else if (window.attachEvent)
	window.attachEvent("onload", setupPanels)
else
	setTimeout( "setupPanels()", 1.0*1000 );

//-->
</script>
<%
END SUB







SUB makeJSImageList( sName, aList, nList )

%>
<script type="text/javascript">

<%

	Response.Write "var g_n" & sName & " = " & nList & ";" & vbCRLF
	Response.Write "var g_a" & sName & " = new Array(" & nList & ");" & vbCRLF

	DIM xFile
	DIM	i
	i = 0
	FOR EACH xFile IN aList
		Response.Write "g_a" & sName & "[" & i & "] = """ & xFile & """;" & vbCRLF
		i = i + 1
	NEXT 'xFile
%>
</script>
<%

END SUB




SUB makePanelStyles()

%>



<link rel="stylesheet" href="<%=g_sTheme%>style.css" type="text/css">
<link rel="stylesheet" href="<%=g_sTheme%>gallery.css" type="text/css">
<style type="text/css">




.thumbnail
{
	position: absolute;
	text-align: center;
	height: <%=g_nThumbSize - g_nThumbMargin%>px;
	width: <%=g_nThumbSize - g_nThumbMargin%>px;
}

.thumbnail a:active img
{
	/*filter: alpha(opacity=30); 
	-moz-opacity:0.30; 
	opacity:0.30;*/
}

.thumbnail a:active
{
}


.thumbSelected
{
	position: absolute;
	text-align: center;
	height: <%=g_nThumbSize - g_nThumbMargin%>px;
	width: <%=g_nThumbSize - g_nThumbMargin%>px;
	border: 1px goldenrod solid;
}

.thumbSelected img
{
	filter: alpha(opacity=30); 
	-moz-opacity:0.30; 
	opacity:0.30;
}


</style>
<!--[if gt IE 6]>
<style type="text/css">

#DisplayPanel
{
	filter: blendTrans(Duration=1.0);
}

</style>
<![endif]-->
<!--[if lt IE 7]>
<style type="text/css">

#DisplayPanel
{
	filter: progid:DXImageTransform.Microsoft.BlendTrans(duration=1.0);
}

</style>
<![endif]-->
<%

END SUB



IF FALSE THEN
%>
</head>
<body>
<%
END IF




SUB makeGalleryBody

%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<tr>
<td valign="top">
<div id="ThumbsFrame">
<div id="ThumbsControlPanel"><a id="ThumbsPrev" href="javascript:thumbsprev()"><img border="0" alt="Previous" src="images/arrow-left.gif" height="19" width="19"></a><a id="ThumbsNext" href="javascript:thumbsnext()"><img border="0" alt="Next" src="images/arrow-right.gif" height="19" width="19"></a></div>
<div id="ThumbsPanel">Loading</div>
</div>
</td>
<td><div id="DisplayPanel">
	<img alt="" src="<%=g_sTheme%>logo-200.gif" width="200" height="200"><br>
	Click on Thumbnail image<br>
	to the left for larger view</div></td>
</tr>
</table>


<%

END SUB


IF FALSE THEN
%>
</body>
</html>

<%
END IF



%>