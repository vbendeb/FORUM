<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<%
'#################################################################################
'## Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or any later version.
'##
'## All copyright notices regarding Snitz Forums 2000
'## must remain intact in the scripts and in the outputted HTML
'## The "powered by" text/logo with a link back to
'## http://forum.snitz.com in the footer of the pages MUST
'## remain visible when the pages are viewed on the internet or intranet.
'##
'## This program is distributed in the hope that it will be useful,
'## but WITHOUT ANY WARRANTY; without even the implied warranty of
'## MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'## GNU General Public License for more details.
'##
'## You should have received a copy of the GNU General Public License
'## along with this program; if not, write to the Free Software
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
'##
'## Support can be obtained from support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## reinhold@bigfoot.com
'##
'## or
'##
'## Snitz Communications
'## C/O: Michael Anderson
'## PO Box 200
'## Harpswell, ME 04079
'#################################################################################

Response.Write	"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"<table width=""95%"" align=""center"" border=""0"" bgcolor=""" & strForumCellColor & """ cellpadding=""0"" cellspacing=""1"">" & vbNewLine & _
		"  <tr>" & vbNewLine & _
		"    <td>" & vbNewLine & _
		"      <table border=""0"" width=""100%"" align=""center"" cellpadding=""4"" cellspacing=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strForumCellColor & """ align=""left"" valign=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & strForumTitle & "</font></td>" & vbNewLine

if request.servervariables("url") = "/FORUM/default.asp" or _
   request.servervariables("url") = "/FORUM/" or _
   request.servervariables("url") = "/FORUM/Default.asp"then
		Response.Write "<td align=""right"" valign=""top"" nowrap>" & _
			"<a href=""http://www2.clustrmaps.com/counter/maps.php?url=http://www.moct.org/FORUM"" id=""clustrMapsLink"">" & _
			"<img src=""http://www2.clustrmaps.com/counter/index2.php?url=http://www.moct.org/FORUM"" " & _
			"style=""border:1px solid;"" alt=""������ ��� ��������"" title=""������ ��� ��������"" id=""clustrMapsImg"" " & _
			"onError=""this.onError=null; this.src='http://clustrmaps.com/images/clustrmaps-back-soon.jpg'; document.getElementById('clustrMapsLink').href='http://clustrmaps.com'"" /></a>" & _
		              "</td>" & vbNewLine
end if
Response.Write	_
		"          <td bgcolor=""" & strForumCellColor & """ align=""right"" valign=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>&copy; " & strCopyright & "</font></td>" & vbNewLine & _
		"          <td bgcolor=""" & strForumCellColor & """ width=""10"" nowrap><a href=""#top""" & dWStatus("Go To Top Of Page...") & " tabindex=""-1"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"<table border=""0"" width=""95%"" align=""center"" cellpadding=""4"" cellspacing=""0"">" & vbNewLine & _
		"  <tr valign=""top"">" & vbNewLine
if strShowTimer = "1" then
	Response.Write	"    <td align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & chkString(replace(strTimerPhrase, "[TIMER]", abs(round(StopTimer(1), 2)), 1, -1, 1),"display") & "</font></td>" & vbNewLine
end if
Response.Write	"    <td align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>"


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<a href=""http://forum.snitz.com"" target=""_blank"" tabindex=""-1""><acronym title=""Powered By: " & strVersion & """>"
if strShowImagePoweredBy = "1" then 
	Response.Write	getCurrentIcon("logo_powered_by.gif||","Powered By: " & strVersion,"")
else
	Response.Write	"Snitz Forums 2000"
end if
Response.Write	"</acronym></a></font></td>" & vbNewline
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT


Response.Write	"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"</font>" & vbNewLine & _
		"<script src=""http://www.google-analytics.com/urchin.js"" type=""text/javascript""></script>" & vbNewLine & _
		"<script type=""text/javascript"">" & vbNewLine & _
		"_uacct = ""UA-1670907-1"";" & vbNewLine & _
		"urchinTracker();</script>" & vbNewLine & _
		"</body>" & vbNewLine & _
		"</html>" & vbNewLine

my_Conn.Close
set my_Conn = nothing 
%>