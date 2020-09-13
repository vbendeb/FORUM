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
on error resume next
strName = my_Conn.Properties(0).name
strValue = my_Conn.Properties(0).value
on error goto 0

if Err.Number <> 0 then
	blnDisplay = False
else
	blnDisplay = True
end if

Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Server&nbsp;Information<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><br /><b>NOTE:</b> The following table will show you values of interest in setting up these forums. Most useful will be the line that shows the APPL_PHYSICAL_PATH. This can be used to properly write your DSN'less Connection String.</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <table border=""0"" align=""center"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" cellspacing=""1"" cellpadding=""1"" align=""center"" width=""100%"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Variable&nbsp;Name</b></font></td>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Value</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine
for each key in Request.ServerVariables
	Response.Write	"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & key & "</b></font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
	if Request.ServerVariables(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write Request.Servervariables(key)
	end if 
	Response.Write	"</font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine
next
if blnDisplay = True then
	'## Code below added to show general ADO/Database Information
	Response.Write	"              <tr>" & vbNewLine & _
			"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Database Connection Properties</font></b></td>" & vbNewLine & _
			"              </tr>" & vbNewLine
	for each item in my_Conn.Properties
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & item.name & "</b></font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><font face=""courier"" size=""" & strDefaultFontSize & """>"
		if item.value = "" then
			Response.Write	"&nbsp;"
		else
			Response.Write	item.value
		end if
		Response.Write	"</font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	next
	'## Code above added to show general ADO/Database Information
end if
Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine
WriteFooter
Response.End
%>
