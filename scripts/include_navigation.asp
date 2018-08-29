<!--#include file="tab_tools.asp"-->
<%







CLASS CTabFormatRed

	PROPERTY GET colorBackground()
		colorBackground = "#990000"
	END PROPERTY
	
	PROPERTY GET colorTab()
		colorTab = "#CC6666"
	END PROPERTY
	
	PROPERTY GET classTabGroup()
		classTabGroup = "TopNavigation"
	END PROPERTY
	
	PROPERTY GET classTabGroupInverted()
		classTabGroupInverted = "TopNavigationInverted"
	END PROPERTY
	
	PROPERTY GET classTab()
		classTab = "navigationtab"
	END PROPERTY
	
	PROPERTY GET classSelected()
		classSelected = "SelectedTab"
	END PROPERTY
	
	PROPERTY GET colorTabSelected()
		colorTabSelected = "#FFCC00"
	END PROPERTY
	
	PROPERTY GET alignTabHoriz()
		alignTabHoriz = "center"
	END PROPERTY
	
	PROPERTY GET imageTL()
		imageTL = "images/pie_tl_red.gif"
	END PROPERTY
	
	PROPERTY GET imageTR()
		imageTR = "images/pie_tr_red.gif"
	END PROPERTY
	
	PROPERTY GET imageBL()
		imageBL = "images/pie_bl_red.gif"
	END PROPERTY

END CLASS





SUB outputPageHeader
%>
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td align="right">
<img alt="Kodiak Photography" src="<%=g_sTheme%>logo.gif"></td>
		<td>
		<div class="clsHeader">
		<span class="clsKodiak">
						Kodiak </span>
		<span class="clsPhoto">
			Photography</span>
		<div class="clsBCG">
			a division of Bear Consulting Group</div></div>
		</td>
<%
IF FALSE THEN
%>
		<td>&nbsp;&nbsp;</td>
		<td class="clsAddress">122 Rachel Drive<br>Huntsville, AL 35806<br>256-722-9128<br><br>
		<span class="pobox">in fo*Ko di ak Ph ot o gr aphy:c om</span></td>
<%
END IF
%>
	</tr>
</table>


<%
END SUB


SUB outputTabs
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td valign="bottom">

<%

	DIM	oTabGen
	DIM	oTabData
	DIM	oTabFormat
	
	SET oTabGen = NEW CTabGenerate
	SET oTabData = NEW CTabDataFiles
	SET oTabFormat = NEW CTabFormatRed
	
	oTabData.Data = g_navtab_aFileList
	
	SET oTabGen.TabData = oTabData
	SET oTabGen.TabFormat = oTabFormat
	oTabGen.MaxCols = 100
	oTabGen.Current = g_navtab_nIndex
	oTabGen.makeTabs
	
	
	SET oTabFormat = Nothing
	SET oTabData = Nothing
	SET oTabGen = Nothing


%>
    </td>
  </tr>
</table>
<%
END SUB



SUB outputTimeBlock
%>
<table border="0" width="100%" cellspacing="0" cellpadding="0">
	<tr>
		<td bgcolor="#996633" class="header">
			<a href="./"><img border="0" src="images/bearpawc3.gif" alt="Bear Consulting Group" align="absmiddle" style="margin-left: 5px; filter:progid:DXImageTransform.Microsoft.Shadow(color=#663300,direction=135,strength=5);"></a>
&nbsp;Bear Consulting Group</td>
	</tr>
</table>
<table border="0" cellspacing="0" width="100%">
  <tr>
    <td bgcolor="#FFCC00"><font size="1">&nbsp; 122 Rachel Drive &nbsp;-&nbsp; Huntsville, AL 35806</font></td>
    <td bgcolor="#FFCC00" align="right"><font size="1">256-722-9128 &nbsp;</font></td>
  </tr>
</table>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" bgcolor="#000000" height="2"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
END SUB






SUB outputPad2
%>
<table border="0" cellspacing="0" width="100%" cellpadding="0">
  <tr>
    <td width="100%" height="6"><spacer type="block" height="1" width="1"></td>
  </tr>
</table>
<%
END SUB





%>
<div id="headerblock">
<%
outputPageHeader
outputTabs
'outputTimeBlock
%>
</div>
<%
'outputPad
IF Response.Buffer THEN Response.Flush

%>