<%@ LANGUAGE="VBSCRIPT" %>
<%


'   PROJECT HONEY POT ADDRESS DISTRIBUTION SCRIPT
'   For more information visit: http://www.projecthoneypot.org/
'   Copyright (C) 2004-2009, Unspam Technologies, Inc.
'   
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'   
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'   
'   You should have received a copy of the GNU General Public License
'   along with this program; if not, write to the Free Software
'   Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA
'   02111-1307  USA
'   
'   If you choose to modify or redistribute the software, you must
'   completely disconnect it from the Project Honey Pot Service, as
'   specified under the Terms of Service Use. These terms are available
'   here:
'   
'   http://www.projecthoneypot.org/terms_of_service_use.php
'   
'   The required modification to disconnect the software from the
'   Project Honey Pot Service is explained in the comments below. To find the
'   instructions, search for:  *** DISCONNECT INSTRUCTIONS ***
'   
'   Generated On: Tue, 19 May 2009 23:03:04 -0400
'   For Domain: www.kodiakphotography.com
'   
'   

'   *** DISCONNECT INSTRUCTIONS ***
'   
'   You are free to modify or redistribute this software. However, if
'   you do so you must disconnect it from the Project Honey Pot Service.
'   To do this, you must delete the lines of code below located between the
'   *** START CUT HERE *** and *** FINISH CUT HERE *** comments. Under the
'   Terms of Service Use that you agreed to before downloading this software,
'   you may not recreate the deleted lines or modify this software to access
'   or otherwise connect to any Project Honey Pot server.
'   
'   *** START CUT HERE ***
'   
 REQUEST_HOST       = "hpr6.projecthoneypot.org"
 REQUEST_PORT       = "80"
 REQUEST_SCRIPT     = "/cgi/serve.php"
'   
'   *** FINISH CUT HERE ***
'   

 HPOT_TAG1          = "36b72ed811b5ae4e89c21492a755c63e"
 HPOT_TAG2          = "10fbe5a14497d48532f0ac0d82286240"
 HPOT_TAG3          = "e5eff424859078ea08098d130e430833"

 CLASS_STYLE_1      = "kasetrix"
 CLASS_STYLE_2      = "drajorauevul"

 DIV1               = "r4ch6zi"

 VANITY_L1          = "MEMBER OF PROJECT HONEY POT"
 VANITY_L2          = "Spam Harvester Protection Network"
 VANITY_L3          = "provided by Unspam"

 DOC_TYPE1          = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">\n"
 HEAD1              = "<html>\n<head>\n"
 HEAD2              = "<title>http://www.kodiakphotography.com</title>\n</head>\n"
 ROBOT1             = "<meta name=""robots"" content=""noindex"">\n<meta name=""robots"" content=""noarchive"">\n<meta name=""robots"" content=""follow"">\n"
 NOCOLLECT1         = "<meta name=""no-email-collection"" content=""/"">\n"
 TOP1               = "<body>\n<div align=""center"" id=""workable"">\n"
 EMAIL1A            = "<a href=""mailto:"
 EMAIL1B            = """ style=""display: none;"">"
 EMAIL1C            = "</a>"
 EMAIL2A            = "<a href=""mailto:"
 EMAIL2B            = """ style=""display:none;"">"
 EMAIL2C            = "</a>"
 EMAIL3A            = "<a style=""display: none;"" href=""mailto:"
 EMAIL3B            = """>"
 EMAIL3C            = "</a>"
 EMAIL4A            = "<a style=""display:none;"" href=""mailto:"
 EMAIL4B            = """>"
 EMAIL4C            = "</a>"
 EMAIL5A            = "<a href=""mailto:"
 EMAIL5B            = """></a>"
 EMAIL5C            = ".."
 EMAIL6A            = "<span style=""display: none;""><a href=""mailto:"
 EMAIL6B            = """>"
 EMAIL6C            = "</a></span>"
 EMAIL7A            = "<span style=""display:none;""><a href=""mailto:"
 EMAIL7B            = """>"
 EMAIL7C            = "</a></span>"
 EMAIL8A            = "<!-- <a href=""mailto:"
 EMAIL8B            = """>"
 EMAIL8C            = "</a> -->"
 EMAIL9A            = "<div id=""" & DIV1 & """><a href=""mailto:"
 EMAIL9B            = """>"
 EMAIL9C            = "</a></div><br><script language=""JavaScript"" type=""text/javascript"">document.getElementById('" & DIV1 & "').innerHTML='';</script>"
 EMAIL10A           = "<a href=""mailto:"
 EMAIL10B           = """><!-- "
 EMAIL10C           = " --></a>"
 LEGAL1             = ""
 LEGAL2             = "\n"
 STYLE1             = "\n<style>a." & CLASS_STYLE_1 & "{color:#FFF;font:bold 10px arial,sans-serif;text-decoration:none;}</style>"
 VANITY1            = "<table cellspacing=""0""cellpadding=""0""border=""0""style=""background:#999;width:230px;""><tr><td valign=""top""style=""padding: 1px 2px 5px 4px;border-right:solid 1px #CCC;""><span style=""font:bold 30px arial,sans-serif;color:#666;top:0px;position:relative;"">@</span></td><td valign=""top"" align=""left"" style=""padding:3px 0 0 4px;""><a href=""http://www.projecthoneypot.org/"" class=""" & CLASS_STYLE_1 & """>" & VANITY_L1 & "</a><br><a href=""http://www.unspam.com""class=""" & CLASS_STYLE_1 & """>" & VANITY_L2 & "<br>" & VANITY_L3 & "</a></td></tr></table>\n"
 BOTTOM1            = "</div>\n</body>\n</html>\n"


Function getLegalContent()
    getLegalContent = "<table border=""0"" cellpadding=""0"" cellspacing=""0""><tr>\n<td><font face=""courier"">&nbsp; &nbsp; &nbsp;&nbsp; <b><font color=""#FFFFFF"">h</font></b>&nbsp; <br>&nbsp;<br>T&#104;e web&#115;it<br>to you sub<br>&#111;ther term<br>Website &#121;o<br>re&#97;d &#116;hem <br>agent&#115; of <br>th<!-- balloon -->em. T&#104;e <br>n&#111;n-transf<br>&#87;ebsite.<br><br><b><font color=""#FFFFFF"">p</font></b>&nbsp; &nbsp; &nbsp; &nbsp;&nbsp; <br>&nbsp;<br>Special re<br>Non-H&#117;man <br>spiders,<!-- trite thievish ritardando incandescent -->&nbsp;b<br>&#112;rograms<font color=""#FFFFFF"">s</font>d<br>aut<!-- herbaceous large spoof essence numb -->omatica<br><br>Email &#97;ddr<br>It is r<!-- long -->&#101;co<br>alone. You<br>has<font color=""#FFFFFF"">s</font>a<font color=""#FFFFFF"">o</font>valu<br>sto&#114;age, &#97;<br>va&#108;ue &#111;f t<br>storing<font color=""#FFFFFF"">s</font>th<br>agreemen&#116; <br><br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <br>&nbsp;<br>Eac&#104; &#112;&#97;rty<br>agai&#110;&#115;t<font color=""#FFFFFF"">f</font>&#116;&#104;<br>(""J<!-- sliest professional tradein -->u<!-- dummy drummers -->d<!-- decadent -->icial<br>the<font color=""#FFFFFF"">d</font>regist<br>suc&#104; law<!-- pale nudists -->&#115; <br>a&#110;d perfor<br>of &#102;ederal<br>&#97;&#110;y a&#99;ti&#111;n<br>Service&#46; Y<br>the a&#98;ove <br><br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <br>&nbsp;<br>You co&#110;sen<br>may app&#101;ar<br>abuse. The<br>Vi&#115;itors &#97;<br><br>VIS&#73;&#84;ORS A<br>PARTY OR S<br>SUBSE&#81;U&#69;NT<br></font></td>\n<td><font face=""courier""><b><font color=""#FFFFFF"">i</font></b>&nbsp; <b><font color=""#FFFFFF"">p</font></b>&nbsp; &nbsp; &nbsp; <br><br>e from wh&#105;<br>&#106;ec&#116; to th<br>s go&#118;ernin<br>u acce&#112;t &#116;<br>ca<!-- birthday billion living -->r&#101;fully&#46;<br>the indivi<br>access<font color=""#FFFFFF"">s</font>rig<br>erable wit<br><br><br>&nbsp; &nbsp;&nbsp; <b>S&#80;&#69;&#67;I</b><br><br>s&#116;rictions<br>Visitor&#115;. <br>ots, index<br>&#101;&#115;igned &#116;o<br>lly.<br><br>&#101;sses<font color=""#FFFFFF"">p</font>&#111;n t<br>gnized tha<br>&nbsp;a&#99;knowled<br>e not less<br>nd/or<font color=""#FFFFFF"">p</font>dist<br>hese &#97;ddre<br>i<!-- criterion womanish ruthless beneficial -->s Website<br>an&#100; expres<br><br>&nbsp; &nbsp;&nbsp; <b><font color=""#FFFFFF"">o</font></b>&nbsp; &nbsp; <br><br>&nbsp;agre&#101;s<font color=""#FFFFFF"">f</font>th<br>&#101; ot&#104;&#101;r in<br>&nbsp;A&#99;t&#105;on<!-- orientation leaved essential many -->"") <br>ere&#100; &#65;dmin<br>are a&#112;p&#108;ie<br>med enti&#114;&#101;<br>&nbsp;and state<br>&nbsp;brought a<br>ou consent<br>agreement.<br><br>&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <br><br>t to havin<br>&nbsp;so&#109;ewhere<br>&nbsp;Identifi&#101;<br>gree not t<br><br>GREE T&#72;&#65;T <br>EN&#68;ING<font color=""#FFFFFF"">d</font>ANY<br><font color=""#FFFFFF"">d</font>B&#82;EACH OF<br></font></td>\n<td><font face=""courier"">&nbsp; &nbsp; &nbsp;&nbsp; <b>&#84;ER</b><br><br>c&#104; you<font color=""#FFFFFF"">t</font>acc<br>e followin<br>&#103; acce&#115;s t<br>&#104;ese te&#114;ms<br>&nbsp;Any No&#110;-H<br>dual(s) wh<br>h&#116;s gra&#110;t&#101;<br>hout the e<br><br><br><b>AL</b>&nbsp;<b>LICENSE</b><br><br>&nbsp;on a vi&#115;i<br>Non-Hum&#97;n <br>e&#114;s, &#114;&#111;bot<br><font color=""#FFFFFF"">h</font>a<!-- professional contagious theatre lowly tubate -->ccess, r<br><br><br>hi&#115; site a<br>t thes&#101; em<br>ge &#97;&#110;d<font color=""#FFFFFF"">p</font>agr<br>&nbsp;t&#104;&#97;n US $<br>ribution o<br>sses. Inte<br>'s email a<br>sly prohi&#98;<br><br>&nbsp; &nbsp;&nbsp; <b>APPLI</b><br><br>at &#97;ny sui<br>&nbsp;connect<!-- seat -->io<br>&#115;hal&#108; be<!-- map alleged -->&nbsp;g<br>istrat&#105;ve <br>d to agree<br>ly<font color=""#FFFFFF"">h</font>within <br><font color=""#FFFFFF"">c</font>courts wi<br>gainst him<br>&nbsp;to e&#108;ectr<br><br><br>&nbsp; &nbsp; <b>REC&#79;RD</b><br><br>g<font color=""#FFFFFF"">h</font>you&#114; Int<br><font color=""#FFFFFF"">t</font>on<font color=""#FFFFFF"">i</font>t&#104;&#105;s p<br>r is &#117;n&#105;qu<br>o us&#101; this<br><br>HAR&#86;EST&#73;NG<br>&nbsp;M&#69;S&#83;A&#71;E(S<br>&nbsp;THESE<font color=""#FFFFFF"">h</font>TE&#82;<br></font></td>\n<td><font face=""courier""><b>MS</b>&nbsp;<b>AND</b>&nbsp;<b>&#67;ON</b><br><br>ess<!-- force -->ed this<br>g condi<!-- rear ulterior -->t&#105;o<br>o the &#87;eb&#115;<br>&nbsp;an&#100; condi<br>&#117;&#109;an<!-- paternity universal roof factory jeans -->&nbsp;Visit<br>o con&#116;rols<br>&#100; to you u<br>xpress wri<br><br><br><b><font color=""#FFFFFF"">k</font>&#82;ESTRICTI</b><br><br>tor's l&#105;ce<br>Vi&#115;itors i<br>s, crawl&#101;r<br>ead, c<!-- classroom grouchy -->ompi<br><br><br>re conside<br>ail add&#114;es<br>ee<font color=""#FFFFFF"">f</font>that ea<br>5&#48;.<font color=""#FFFFFF"">s</font>You fu<br>&#102; the&#115;e<font color=""#FFFFFF"">a</font>&#97;d<br>nti&#111;&#110;al co<br>d&#100;&#114;esses i<br>ited.<br><br><b>C&#65;&#66;LE</b>&nbsp;<b>LA&#87;</b>&nbsp;<br><br>t, action <br>n with<font color=""#FFFFFF"">d</font>or <br>overne&#100; by<br>Conta&#99;t (t<br>m&#101;n&#116;s be<!-- round squalid ruin fellow -->&#116;w<br>the Admin<font color=""#FFFFFF"">p</font><br>t&#104;&#105;n<font color=""#FFFFFF"">k</font>t&#104;e &#65;<br>&nbsp;in<font color=""#FFFFFF"">e</font>&#99;o&#110;nec<br>on&#105;c s&#101;&#114;v&#105;<br><br><br><b>S</b>&nbsp;<b>OF<font color=""#FFFFFF"">p</font>V&#73;SI&#84;</b><br><br>&#101;&#114;net Pr<!-- lengthy forward -->o&#116;<br>age (&#116;he<font color=""#FFFFFF"">c</font>""<br>ely matche<br>&nbsp;&#97;ddre&#115;s f<br><br>, GATHE&#82;IN<br>) TO THE &#73;<br>&#77;S<font color=""#FFFFFF"">k</font>O&#70; SERV<br></font></td>\n<td><font face=""courier""><b>DITIONS</b>&nbsp;<b>OF</b><br><br>&nbsp;agreement<br>ns. &#84;hese <br>it&#101;&#46; By vi<br>ti&#111;ns<font color=""#FFFFFF"">o</font>(the<br>o&#114;s to th&#101;<br>, authors<!-- behaviour strict utter -->&nbsp;<br>nder the T<br>tt&#101;n pe&#114;mi<br><br><br><b>O&#78;S</b>&nbsp;<b>&#70;OR</b>&nbsp;<b>&#78;O</b><br><br>nse t&#111; ac&#99;<br>nclude, bu<br>s<!-- jaunty dusty continuing tooth undesirable -->, &#104;arvest<br>le &#111;r g&#97;&#116;h<br><br><br>red propri<br>ses are pr<br>c<!-- meteoric -->h email a<br>&#114;th<!-- keeper woodpeckers realistic relational -->er agre<br>dres&#115;es su<br>l&#108;ection&#44; <br>s recogniz<br><br><br><b>AND</b>&nbsp;<b>JURISD</b><br><br>or &#112;roc&#101;ed<br>ar&#105;sing fr<br>&nbsp;&#116;he &#108;aw o<br>he ""A&#100;min <br>een Adm&#105;n<font color=""#FFFFFF"">c</font><br>State. &#89;o&#117;<br>dmin S<!-- bat binomial -->tat&#101;<br>tion<font color=""#FFFFFF"">i</font>with <br>ce of proc<br><br><br><b>OR</b>&nbsp;<b>U&#83;&#69;</b>&nbsp;<b>A&#78;D</b><br><br>&#111;col addr<!-- chamber conic -->e<br>Iden&#116;ifier<br>&#100; t&#111; your <br>or<font color=""#FFFFFF"">e</font>&#97;ny rea<br><br>G, ST&#79;RING<br>DENTIFI&#69;R <br>IC&#69;.<br></font></td>\n<td><font face=""courier"">&nbsp;<b>USE<font color=""#FFFFFF"">g</font></b><br><br><font color=""#FFFFFF"">k</font>(""th&#101; Web<br>te&#114;m&#115; are <br>siting (in<br>&nbsp;""&#84;erm&#115; of<br>&nbsp;Website s<br>or<!-- magical intervention particle metal great -->&nbsp;otherwi<br>e&#114;ms of Se<br>ssion of t<br><br><br><b>N-HUMAN</b>&nbsp;<b>&#86;&#73;</b><br><br>es&#115; the We<br>t a&#114;e n&#111;t <br>ers, or an<br>er content<br><br><br>etary inte<br>&#111;vided<font color=""#FFFFFF"">p</font>&#102;or<br>dd&#114;ess th&#101;<br>e that the<br>b&#115;tantial&#108;<br>harves&#116;ing<br>&#101;d as &#97; vi<br><br><br><b>ICTION</b>&nbsp;<br><br>ing<font color=""#FFFFFF"">c</font>b&#114;ough<br>om &#116;h&#101; Ter<br>f the stat<br>State""<!-- soil civic -->) fo<br>Sta&#116;e resi<br>&nbsp;consent<font color=""#FFFFFF"">o</font>t<br>. &#89;ou cons<br>&#98;reaches o<br>ess regard<br><br><br>&nbsp;<b>ABUSE</b>&nbsp;<br><br>ss recorde<br>"") if &#119;e s<br>In&#116;erne&#116;<font color=""#FFFFFF"">h</font>P<br>son.<br><br>, T&#82;ANSFER<br>CONS&#84;IT&#85;TE<br><br></font></td>\n<td><font face=""courier""><br><br>site"") is <br>in addit&#105;o<br>&nbsp;any manne<br>&nbsp;&#83;ervice"")<br>hal&#108; be co<br>se mak&#101;s u<br>&#114;vic&#101; are<br>he &#111;wner o<br><br><br><b>SI&#84;ORS</b>&nbsp;<br><br>bsite appl<br>li&#109;i<!-- sensitivity endothermic -->ted to<br>&#121; other c&#111;<br>&nbsp;from the<font color=""#FFFFFF"">f</font><br><br><br>l&#108;e&#99;tual p<br><font color=""#FFFFFF"">i</font>hu<!-- sparrows wretched motley top -->man vis<br>&nbsp;Website c<br>&nbsp;c&#111;mp&#105;lati<br>y diminish<br>&#44;<font color=""#FFFFFF"">e</font>gathe&#114;&#105;&#110;<br>olatio&#110; of<br><br><br><br><br>t &#98;y such <br>ms of &#83;e&#114;&#118;<br>e of resid<br>r th&#101; Web&#115;<br>&#100;en&#116;s<font color=""#FFFFFF"">f</font>en&#116;e<br>o the juri<br>ent to<font color=""#FFFFFF"">p</font>the<br>f<!-- such viewpoint -->&nbsp;these T&#101;<br>ing &#97;ction<br><br><br><br><br>d.<font color=""#FFFFFF"">e</font>An<font color=""#FFFFFF"">g</font>em&#97;i<br>us&#112;ect pot<br>ro&#116;ocol ad<br><br><br>RI&#78;G TO A <br>S AN &#65;C&#67;&#69;P<br><br></font></td>\n<td><font face=""courier""><br><br>provi&#100;ed<br>n to &#97;ny<br>r&#41; &#116;he<br>. Plea&#115;e<br>&#110;sidered<br>se<font color=""#FFFFFF"">f</font>of<br><br>&#102; the<br><br><br><br><br>y t&#111;<br>, web<br>mput&#101;r<br>Website<br><br><br>r&#111;perty.<br>itors<br>onta&#105;ns<br>on&#44;<br>e&#115;<font color=""#FFFFFF"">d</font>the<br>g,<font color=""#FFFFFF"">g</font>and/or<br><font color=""#FFFFFF"">i</font>this<br><br><br><br><br>p&#97;&#114;&#116;&#121;<br>ic&#101;<br>&#101;nce of<br>ite<font color=""#FFFFFF"">c</font>a&#115;<br>red in&#116;o<br>s<!-- graphics -->&#100;i<!-- impacted straightforward able islands christian -->c&#116;ion<br>&nbsp;venu&#101; in<br>rms &#111;f<br>s under<br><br><br><br><br>l<font color=""#FFFFFF"">g</font>address<br>ential<br>dress.<br><br><br>THIRD<br>TANCE<font color=""#FFFFFF"">h</font>A&#78;D<br><br></font></td>\n</tr>\n</table>\n<br>"
End Function




Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32

Private m_lOnBits(30)
Private m_l2Power(30)
 
    m_lOnBits(0) = CLng(1)
    m_lOnBits(1) = CLng(3)
    m_lOnBits(2) = CLng(7)
    m_lOnBits(3) = CLng(15)
    m_lOnBits(4) = CLng(31)
    m_lOnBits(5) = CLng(63)
    m_lOnBits(6) = CLng(127)
    m_lOnBits(7) = CLng(255)
    m_lOnBits(8) = CLng(511)
    m_lOnBits(9) = CLng(1023)
    m_lOnBits(10) = CLng(2047)
    m_lOnBits(11) = CLng(4095)
    m_lOnBits(12) = CLng(8191)
    m_lOnBits(13) = CLng(16383)
    m_lOnBits(14) = CLng(32767)
    m_lOnBits(15) = CLng(65535)
    m_lOnBits(16) = CLng(131071)
    m_lOnBits(17) = CLng(262143)
    m_lOnBits(18) = CLng(524287)
    m_lOnBits(19) = CLng(1048575)
    m_lOnBits(20) = CLng(2097151)
    m_lOnBits(21) = CLng(4194303)
    m_lOnBits(22) = CLng(8388607)
    m_lOnBits(23) = CLng(16777215)
    m_lOnBits(24) = CLng(33554431)
    m_lOnBits(25) = CLng(67108863)
    m_lOnBits(26) = CLng(134217727)
    m_lOnBits(27) = CLng(268435455)
    m_lOnBits(28) = CLng(536870911)
    m_lOnBits(29) = CLng(1073741823)
    m_lOnBits(30) = CLng(2147483647)
    
    m_l2Power(0) = CLng(1)
    m_l2Power(1) = CLng(2)
    m_l2Power(2) = CLng(4)
    m_l2Power(3) = CLng(8)
    m_l2Power(4) = CLng(16)
    m_l2Power(5) = CLng(32)
    m_l2Power(6) = CLng(64)
    m_l2Power(7) = CLng(128)
    m_l2Power(8) = CLng(256)
    m_l2Power(9) = CLng(512)
    m_l2Power(10) = CLng(1024)
    m_l2Power(11) = CLng(2048)
    m_l2Power(12) = CLng(4096)
    m_l2Power(13) = CLng(8192)
    m_l2Power(14) = CLng(16384)
    m_l2Power(15) = CLng(32768)
    m_l2Power(16) = CLng(65536)
    m_l2Power(17) = CLng(131072)
    m_l2Power(18) = CLng(262144)
    m_l2Power(19) = CLng(524288)
    m_l2Power(20) = CLng(1048576)
    m_l2Power(21) = CLng(2097152)
    m_l2Power(22) = CLng(4194304)
    m_l2Power(23) = CLng(8388608)
    m_l2Power(24) = CLng(16777216)
    m_l2Power(25) = CLng(33554432)
    m_l2Power(26) = CLng(67108864)
    m_l2Power(27) = CLng(134217728)
    m_l2Power(28) = CLng(268435456)
    m_l2Power(29) = CLng(536870912)
    m_l2Power(30) = CLng(1073741824)

Private Function LShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        LShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And 1 Then
            LShift = &H80000000
        Else
            LShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If

    If (lValue And m_l2Power(31 - iShiftBits)) Then
        LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
    Else
        LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
    End If
End Function

Private Function RShift(lValue, iShiftBits)
    If iShiftBits = 0 Then
        RShift = lValue
        Exit Function
    ElseIf iShiftBits = 31 Then
        If lValue And &H80000000 Then
            RShift = 1
        Else
            RShift = 0
        End If
        Exit Function
    ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
        Err.Raise 6
    End If
    
    RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

    If (lValue And &H80000000) Then
        RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
    End If
End Function

Private Function RotateLeft(lValue, iShiftBits)
    RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function

Private Function AddUnsigned(lX, lY)
    Dim lX4
    Dim lY4
    Dim lX8
    Dim lY8
    Dim lResult
 
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000
 
    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
 
    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If
 
    AddUnsigned = lResult
End Function

Private Function F(x, y, z)
    F = (x And y) Or ((Not x) And z)
End Function

Private Function G(x, y, z)
    G = (x And z) Or (y And (Not z))
End Function

Private Function H(x, y, z)
    H = (x Xor y Xor z)
End Function

Private Function I(x, y, z)
    I = (y Xor (x Or (Not z)))
End Function

Private Sub FF(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub GG(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub HH(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Sub II(a, b, c, d, x, s, ac)
    a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac))
    a = RotateLeft(a, s)
    a = AddUnsigned(a, b)
End Sub

Private Function ConvertToWordArray(sMessage)
    Dim lMessageLength
    Dim lNumberOfWords
    Dim lWordArray()
    Dim lBytePosition
    Dim lByteCount
    Dim lWordCount
    
    Const MODULUS_BITS = 512
    Const CONGRUENT_BITS = 448
    
    lMessageLength = Len(sMessage)
    
    lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
    ReDim lWordArray(lNumberOfWords - 1)
    
    lBytePosition = 0
    lByteCount = 0
    Do Until lByteCount >= lMessageLength
        lWordCount = lByteCount \ BYTES_TO_A_WORD
        lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
        lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
        lByteCount = lByteCount + 1
    Loop

    lWordCount = lByteCount \ BYTES_TO_A_WORD
    lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE

    lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)

    lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
    lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
    
    ConvertToWordArray = lWordArray
End Function

Private Function WordToHex(lValue)
    Dim lByte
    Dim lCount
    
    For lCount = 0 To 3
        lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
        WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
    Next
End Function

Public Function MD5(sMessage)
    Dim x
    Dim k
    Dim AA
    Dim BB
    Dim CC
    Dim DD
    Dim a
    Dim b
    Dim c
    Dim d
    
    Const S11 = 7
    Const S12 = 12
    Const S13 = 17
    Const S14 = 22
    Const S21 = 5
    Const S22 = 9
    Const S23 = 14
    Const S24 = 20
    Const S31 = 4
    Const S32 = 11
    Const S33 = 16
    Const S34 = 23
    Const S41 = 6
    Const S42 = 10
    Const S43 = 15
    Const S44 = 21

    x = ConvertToWordArray(sMessage)
    
    a = &H67452301
    b = &HEFCDAB89
    c = &H98BADCFE
    d = &H10325476

    For k = 0 To UBound(x) Step 16
        AA = a
        BB = b
        CC = c
        DD = d
    
        FF a, b, c, d, x(k + 0), S11, &HD76AA478
        FF d, a, b, c, x(k + 1), S12, &HE8C7B756
        FF c, d, a, b, x(k + 2), S13, &H242070DB
        FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
        FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
        FF d, a, b, c, x(k + 5), S12, &H4787C62A
        FF c, d, a, b, x(k + 6), S13, &HA8304613
        FF b, c, d, a, x(k + 7), S14, &HFD469501
        FF a, b, c, d, x(k + 8), S11, &H698098D8
        FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
        FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
        FF b, c, d, a, x(k + 11), S14, &H895CD7BE
        FF a, b, c, d, x(k + 12), S11, &H6B901122
        FF d, a, b, c, x(k + 13), S12, &HFD987193
        FF c, d, a, b, x(k + 14), S13, &HA679438E
        FF b, c, d, a, x(k + 15), S14, &H49B40821
    
        GG a, b, c, d, x(k + 1), S21, &HF61E2562
        GG d, a, b, c, x(k + 6), S22, &HC040B340
        GG c, d, a, b, x(k + 11), S23, &H265E5A51
        GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
        GG a, b, c, d, x(k + 5), S21, &HD62F105D
        GG d, a, b, c, x(k + 10), S22, &H2441453
        GG c, d, a, b, x(k + 15), S23, &HD8A1E681
        GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
        GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
        GG d, a, b, c, x(k + 14), S22, &HC33707D6
        GG c, d, a, b, x(k + 3), S23, &HF4D50D87
        GG b, c, d, a, x(k + 8), S24, &H455A14ED
        GG a, b, c, d, x(k + 13), S21, &HA9E3E905
        GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
        GG c, d, a, b, x(k + 7), S23, &H676F02D9
        GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
            
        HH a, b, c, d, x(k + 5), S31, &HFFFA3942
        HH d, a, b, c, x(k + 8), S32, &H8771F681
        HH c, d, a, b, x(k + 11), S33, &H6D9D6122
        HH b, c, d, a, x(k + 14), S34, &HFDE5380C
        HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
        HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
        HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
        HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
        HH a, b, c, d, x(k + 13), S31, &H289B7EC6
        HH d, a, b, c, x(k + 0), S32, &HEAA127FA
        HH c, d, a, b, x(k + 3), S33, &HD4EF3085
        HH b, c, d, a, x(k + 6), S34, &H4881D05
        HH a, b, c, d, x(k + 9), S31, &HD9D4D039
        HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
        HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
        HH b, c, d, a, x(k + 2), S34, &HC4AC5665
    
        II a, b, c, d, x(k + 0), S41, &HF4292244
        II d, a, b, c, x(k + 7), S42, &H432AFF97
        II c, d, a, b, x(k + 14), S43, &HAB9423A7
        II b, c, d, a, x(k + 5), S44, &HFC93A039
        II a, b, c, d, x(k + 12), S41, &H655B59C3
        II d, a, b, c, x(k + 3), S42, &H8F0CCC92
        II c, d, a, b, x(k + 10), S43, &HFFEFF47D
        II b, c, d, a, x(k + 1), S44, &H85845DD1
        II a, b, c, d, x(k + 8), S41, &H6FA87E4F
        II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
        II c, d, a, b, x(k + 6), S43, &HA3014314
        II b, c, d, a, x(k + 13), S44, &H4E0811A1
        II a, b, c, d, x(k + 4), S41, &HF7537E82
        II d, a, b, c, x(k + 11), S42, &HBD3AF235
        II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
        II b, c, d, a, x(k + 9), S44, &HEB86D391
    
        a = AddUnsigned(a, AA)
        b = AddUnsigned(b, BB)
        c = AddUnsigned(c, CC)
        d = AddUnsigned(d, DD)
    Next
    
    MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
                    
Function getFileContents(ByRef Filepath)

   Const ForReading = 1
   Const TristateUseDefault = -2
   
   Dim FSO
   set FSO = server.createObject("Scripting.FileSystemObject")
   
   if FSO.FileExists(Filepath) Then
      Set TextStream = FSO.OpenTextFile(Filepath, ForReading, False, TristateUseDefault)
      Dim Contents
   	  Contents = TextStream.ReadAll
   	  'Response.write("<pre>" & Contents & "</pre>")
      TextStream.Close
   	  Set TextStream = nothing
   Else
      Response.Write("<font color=""red"">WARNING: File " & Filepath & " could not be read!</font>")
      getFileContents = nothing
      exit function 
   
   End If
   
   Set FSO = nothing

   getFileContents = Contents
End Function

Function getDocType() 
	getDocType = DOC_TYPE1
End Function

Function getHeadHTML1() 
	getHeadHTML1 = HEAD1
End Function

Function getRobotHTML() 
	getRobotHTML = ROBOT1
End Function

Function getNoCollectHTML() 
	getNoCollectHTML = NOCOLLECT1
End Function
 
Function getHeadHTML2() 
	getHeadHTML2 = HEAD2
End Function

Function getTopHTML() 
	getTopHTML = TOP1
End Function
 
Function getEmailHTML(Method, Email)
	Select Case Method
		Case 0:
			getEmailHTML = ""
		Case 1:
			getEmailHTML = EMAIL1A & Email & EMAIL1B & Email & EMAIL1C
		Case 2:
			getEmailHTML = EMAIL2A & Email & EMAIL2B & Email & EMAIL2C
		Case 3:
			getEmailHTML = EMAIL3A & Email & EMAIL3B & Email & EMAIL3C
		Case 4:
			getEmailHTML = EMAIL4A & Email & EMAIL4B & Email & EMAIL4C
		Case 5:
			getEmailHTML = EMAIL5A & Email & EMAIL5B & Email & EMAIL5C
		Case 6:
			getEmailHTML = EMAIL6A & Email & EMAIL6B & Email & EMAIL6C
		Case 7:
			getEmailHTML = EMAIL7A & Email & EMAIL7B & Email & EMAIL7C
		Case 8:
			getEmailHTML = EMAIL8A & Email & EMAIL8B & Email & EMAIL8C
		Case 9:
			getEmailHTML = EMAIL9A & Email & EMAIL9B & Email & EMAIL9C
        case Else:
			getEmailHTML = EMAIL10A & Email & EMAIL10B & Email & EMAIL10C
	End Select
End Function

Function getLegalHTML
	getLegalHTML = LEGAL1 & getLegalContent() & LEGAL2
End Function

Function getStyleHTML
	getStyleHTML = STYLE1
End Function

Function getVanityHTML
	getVanityHTML = VANITY1
End Function

Function getBottomHTML
	getBottomHTML = BOTTOM1
End Function

Function performRequest(Request) 
	ResponseStr = ""
	URL = ""
	
	Set srvXmlHttp = Server.CreateObject("MICROSOFT.XMLHTTP")
	
	URL = "http://" & REQUEST_HOST & REQUEST_SCRIPT
    
	srvXmlHttp.open "POST", URL, false
	srvXmlHttp.setRequestHeader "Cache-Control", "no-cache"
	srvXmlHttp.setRequestHeader "User-Agent", "PHPot " & HPOT_TAG2
	srvXmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	srvXmlHttp.setRequestHeader "Connection", "close"
    srvXmlHttp.send Request
    
    performRequest = srvXmlHttp.responseText
End Function

Function prepareRequest() 
	Set postvars = CreateObject("Scripting.Dictionary")

	postvars.Add "tag1", HPOT_TAG1
	postvars.Add "tag2", HPOT_TAG2
	postvars.Add "tag3", HPOT_TAG3
	postvars.Add "tag4", md5(getFileContents(Request.ServerVariables("PATH_TRANSLATED")))
	postvars.Add "ip", Server.URLEncode(Request.ServerVariables("REMOTE_ADDR"))
	postvars.Add "svrn", Server.URLEncode(Request.ServerVariables("SERVER_NAME"))
	postvars.Add "svp", Server.URLEncode(Request.ServerVariables("SERVER_PORT"))
	postvars.Add "svip", Server.URLEncode(Request.ServerVariables("SERVER_ADDR"))
	postvars.Add "rquri", Server.URLEncode(Request.ServerVariables("URL"))
	postvars.Add "sn", Replace(Server.URLEncode(Request.ServerVariables("SCRIPT_NAME")), " ", "%20")
	postvars.Add "ref", Server.URLEncode(Request.ServerVariables("HTTP_REFERER"))
	postvars.Add "uagnt", Server.URLEncode(Request.ServerVariables("HTTP_USER_AGENT"))
	
	Set prepareRequest = postvars
End Function

Function transcribeResponse(ByVal response)
	Set settings = CreateObject("Scripting.Dictionary")
	Arr = Split(URLDecode(response), Chr(10))
	isParam = false

    
    For j = 0 to UBound(Arr)
		If Arr(j) = "<END>" Then isParam = false
		
		If isParam Then
			pieces = Split(Arr(j), "=", 2)
			If UBound(pieces) = 1 Then
				settings.Add pieces(0), pieces(1)
			End If
		End If
		
		If Arr(j) = "<BEGIN>" Then isParam = true
		
    Next
    
    
    If settings.Exists("directives") Then
		settings.Item("directives") = Split(settings.Item("directives"), ",")
    End If
    
    Set transcribeResponse = settings
End Function

Function URLDecode(ByRef str)	
	Set re = New RegExp
	str = Replace(str, "+", " ")
	re.Pattern = "%([0-9a-fA-F]{2})"
	re.Global = True
	URLDecode = re.Replace(str, GetRef("URLDecodeHex"))
End Function

Function URLDecodeHex(match, hex_digits, pos, source)	
	URLDecodeHex = chr("&H" & hex_digits)
End Function

Function formatHTML(ByRef str)
	formatHTML = Replace(str, "\n", VBCrLf)
End Function

Function echo(ByRef str)
	Response.Write(formatHTML(str))
End Function
    
RequestText = ""
ResponseText = ""
Set Post = prepareRequest

Items = Post.Items
Keys = Post.Keys
For j = 0 to Post.Count -1
	RequestText = RequestText & "&" & Keys(j) & "=" & Items(j)
Next

RequestText = Mid(RequestText, 2)
ResponseText = performRequest(RequestText)
Set settings = transcribeResponse(ResponseText)

directives = settings.Item("directives")
email = settings.Item("email")
emailmethod = settings.Item("emailmethod")

Response.AddHeader "Cache-Control", "no-cache"


%>

<% If directives(0) And directives(0) = "1" Then echo(getDocType)%>
<% If settings("injDocType") Then echo(settings("injDocTypeMsg"))%>
<% If directives(1) And directives(1) = "1" Then echo(getHeadHTML1)%>
<% If settings("injHead1HTML") Then echo(settings("injHead1HTMLMsg"))%>
<% If directives(8) And directives(8) = "1" Then echo(getRobotHTML)%>
<% If settings("injRobotHTML") Then echo(settings("injRobotHTMLMsg"))%>
<% If directives(9) And directives(9) = "1" Then echo(getNoCollectHTML)%>
<% If settings("injNoCollectHTML") Then echo(settings("injNoCollectHTMLMsg"))%>
<% If directives(1) And directives(1) = "1" Then echo(getHeadHTML2)%>
<% If settings("injHead2HTML") Then echo(settings("injHead2HTMLMsg"))%>
<% If directives(2) And directives(2) = "1" Then echo(getTopHTML)%>
<% If settings("injTopHTML") Then echo(settings("injTopHTMLMsg"))%>
<%
   IF settings("actMsgOn") <> "" Then echo(settings("actMsg"))
    
   IF settings("errMsgOn") <> "" Then echo(settings("errMsg"))

   IF settings("customMsgOn") <> "" Then echo(settings("customMsg"))
%>
<% If directives(3) And directives(3) = "1" Then echo(getLegalHTML)%>
<% If settings("injLegalHTML") Then echo(settings("injLegalHTMLMsg"))%>
<%
   IF settings("altLegalOn") <> "" Then echo(settings("altLegalMsg"))
%>

<% If directives(4) And directives(4) = "1" Then echo(getEmailHTML(emailmethod, email))%>
<% If settings("injEmailHTML") Then echo(settings("injEmailHTMLMsg"))%>
<% If directives(5) And directives(5) = "1" Then echo(getStyleHTML)%>
<% If settings("injStyleHTML") Then echo(settings("injStyleHTMLMsg"))%>
<% If directives(6) And directives(6) = "1" Then echo(getVanityHTML)%>
<% If settings("injVanityHTML") Then echo(settings("injVanityHTMLMsg"))%>
<%
   IF settings("altVanityOn") <> "" Then echo(settings("altVanityMsg"))
%>
<% If directives(7) And directives(7) = "1" Then echo(getBottomHTML)%>
<% If settings("injBottomHTML") Then echo(settings("injBottomHTMLMsg"))%>
