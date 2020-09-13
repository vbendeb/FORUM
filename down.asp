<%@CODEPAGE=1251 %>
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
%>
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE file="inc_func_common.asp" -->
<%
Dim status, info1, info2, fStatus
mlev = request("mlev")
status = Application(strCookieURL & "down")
fStatus = request.form("status")
DMessage = request.Form("DownMessage")

if DMessage = "" then
	DMessage = Application(strCookieURL & "DownMessage")
end if

if status = "" then
	status = false
end if

if (not isEmpty(fStatus)) and (Session(strCookieURL & "Approval") = "15916941253") then 
	if status then
		Application.lock
		Application(strCookieURL & "down") = false
		Application(strCookieURL & "DownMessage") = ""
		Application.unlock
		status = false
	else
		Application.lock
		Application(strCookieURL & "down") = true
		Application(strCookieURL & "DownMessage") = DMessage
		Application.unlock
		status = true
	end if
end if

if status then
	info1 = "down"
	info2 = "Start"
else
	info1 = "running"
	info2 = "Stop"
end if
if request.form("location") <> "" then response.redirect(request.form("location"))
if Session(strCookieURL & "Approval") = "15916941253" Then
	strScriptName = request.servervariables("script_name")
	Response.Write	"<html>" & vbNewLine & _
			"<head>" & vbNewline & _
			"<title>" & GetNewTitle(strScriptName) & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<meta http-equiv=""Content-Type""; content=""text/html""; charset=""windows-1251"">" & vbNewline

	Response.Write	"</head>" & vbNewLine & _
			"<body background=""" & strPageBGImageURL & """ bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
			"<table border=""0"" width=""100%"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"    " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
			"    " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
			"    " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Forum&nbsp;Maintenance<br /></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"<br />" & vbNewLine & _
			"<form action=""down.asp"" method=""post"">" & vbNewLine & _
			"<table width=""600"" border=""0"" cellspacing=""0"" cellpadding=""10"" align=""center"">" & vbNewLine & _
			"  <tr align=""center"">" & vbNewLine & _
			"    <td><p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Welcome Administrator. The current status of the boards is <font color=""" & strHiLiteFontColor & """>" & info1 & "</font>.</b></font></p>" & vbNewLine & _
			"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>would you like to : </font></p>" & vbNewLine & _
			"    <input type=""submit"" value=""" & info2 & " the board"" name=""Submit"">" & vbNewLine & _
			"    <input type=""hidden"" value=""" & request("target") & """ name=""location"">" & vbNewLine & _
			"    <input type=""hidden"" name=""status"" value=""" & status & """>" & vbNewLine & _
			"    </td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>The message below will appear when the board is closed.</font></p>" & vbNewLine & _
			"    <textarea cols=""80"" rows=""12"" name=""DownMessage"" wrap=""soft"">" & Application(strCookieURL & "DownMessage") & "</textarea></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</form><br />" & vbNewLine
else  
	if mlev = 4 then 
		Response.Redirect "admin_login.asp?target=down.asp?mlev=" & mLev
	elseif not Application(strCookieURL & "down") then 
		response.redirect("default.asp")
	end if

	strScriptName = request.servervariables("script_name")
	Response.Write	"<html>" & vbNewLine & _
			"<head>" & vbNewline & _
			"<title>" & GetNewTitle(strScriptName) & "</title>" & vbNewline


'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write	"<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-02 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & vbNewline 
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<meta http-equiv=""Content-Type""; content=""text/html""; charset=""windows-1251"">" & vbNewline

	Response.Write	"</head>" & vbNewLine & _
			"<body background=""" & strPageBGImageURL & """ bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ link=""" & strLinkColor & """ aLink=""" & strActiveLinkColor & """ vLink=""" & strVisitedLinkColor & """>" & vbNewLine & _
			"<p>&nbsp;</p>" & vbNewLine & _
			"<div align=""center""><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>" & strForumTitle & " is currently closed.</font></p></div>" & vbNewLine & _
			"<table border=""0"" cellspacing=""0"" cellpadding=""10"" align=""center"" width=""50%"">" & vbNewLine & _
			"  <tr align=""center"">" & vbNewLine & _
			"    <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"    The Administrator has chosen to close<br />this forum with the following reason:" & vbNewLine & _
			"    <p><b>" & Application(strCookieURL & "DownMessage") & "</b></p>" & vbNewLine & _
			"    </font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr align=""center"">" & vbNewLine & _
			"    <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""admin_login.asp?target=down.asp"">Administrator Login</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</body>" & vbNewLine & _
			"</html>" & vbNewLine
end if
%>