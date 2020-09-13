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
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" align=center width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Forum&nbsp;Variables&nbsp;Information<br /><br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>NOTE:</b> The following table will show you values of the different variables used by the Forum.</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <table border=""0"" align=""center"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" cellspacing=""1"" cellpadding=""1"" align=""center"" width=""100%"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Variable&nbsp;Name</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Value</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>General&nbsp;information</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strCookieUrl</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(StrCookieUrl, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strUniqueID</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(StrUniqueID, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strAuthType</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(strAuthType, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strDBNTSQLName</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(strDBNTSQLName, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strDBNTUserName</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(strDBNTUserName, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>strDBType</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(strDBType, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>intCookieDuration</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & ChkString(intCookieDuration, "admindisplay") & "</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Cookies</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine
for each key in Request.Cookies 
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		if Request.Cookies(key).HasKeys then
			for each subkey in Request.Cookies(key)
				Response.Write	"              <tr>" & vbNewLine & _
						"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & chkString(key, "admindisplay") & " (" & chkString(subkey, "admindisplay") & ")</b></font></td>" & vbNewLine & _
						"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
				if Request.Cookies(key)(subkey) = "" then
					Response.Write "&nbsp;"
				else
					Response.Write ChkString(CStr(Request.Cookies(key)(subkey)), "admindisplay")
				end if 
				Response.Write	"</font></td>" & vbNewline & _
						"              </tr>" & vbNewline
			next
		else
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & chkString(key, "admindisplay") & "</b></font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
			if Request.Cookies(key) = "" then
				Response.Write	"&nbsp;"
			else
				Response.Write	ChkString(CStr(Request.Cookies(key)), "admindisplay")
			end if 
			Response.Write	"</font></td>" & vbNewline & _
					"              </tr>" & vbNewline
		end if
	end if
next
Response.Write	"              <tr>" & vbNewLine & _
		"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Session&nbsp;variables</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine
for each key in Session.Contents
	if not IsArray(Session.Contents(key)) then
		if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & ChkString(key, "admindisplay") & "</b></font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
			if Session.Contents(key) = "" then
				Response.Write "&nbsp;"
			else
				Response.Write chkString(CStr(Session.Contents(key)), "admindisplay")
			end if 
			Response.Write	"</font></td>" & vbNewline & _
					"              </tr>" & vbNewline
		end if
	end if
next 
Response.Write	"              <tr>" & vbNewLine & _
		"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Application&nbsp;variables</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine
for each key in Application.Contents
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & chkString(key, "admindisplay") & "</b></font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
		if Application.Contents(key) = "" then
			Response.Write	"&nbsp;"
		else
			Response.Write	chkString(CStr(Application.Contents(key)), "admindisplay")
		end if 
		Response.Write	"</font></td>" & vbNewline & _
				"              </tr>" & vbNewline
	end if
next 
Response.Write	"            </table>" & vbNewline & _
		"          </td>" & vbNewline & _
		"        </tr>" & vbNewline & _
		"      </table>" & vbNewline & _
		"      <br />" & vbNewline
WriteFooter
Response.End
%>
