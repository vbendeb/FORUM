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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
fName = strDBNTFUserName
fPassword = ChkString(Request.Form("Password"), "SQLString")

RequestMethod = Request.ServerVariables("Request_method")

if RequestMethod = "POST" Then
	strEncodedPassword = sha256("" & fPassword)

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & fName & "' AND "
	strSql = strSql & "       M_PASSWORD = '" & strEncodedPassword & "' AND "
	strSql = strSql & "       M_LEVEL = 3"
	
	Set dbRs = my_Conn.Execute(strSql)
		
	If not(dbRS.EOF) and ChkQuoteOk(fName) and ChkQuoteOk(fPassword) then 

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Login was successful!</font></p>" & vbNewLine
		Session(strCookieURL & "Approval") = "15916941253"
		Response.Write	"      " & strParagraphFormat1 & "<a href="""
		if Request("target") = "" then
			Response.Write	"admin_home.asp"
		else
			Response.Write	request("target")
		end if
		Response.Write	""">Click here to Continue</a></font></p>" & vbNewLine

		Response.Write	"      <meta http-equiv=""Refresh"" content=""2; URL="
		if Request("target") = "" then
			Response.Write	"admin_home.asp"
		else
			Response.Write	request("target")
		end if
		Response.Write	""">" & vbNewline

		WriteFooterShort
		Response.End

	else
		Response.Write	"      <center>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>There has been a problem!</font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You are not allowed access.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "If you think you have reached this message in error, please try again.</font></p>" & vbNewLine & _
				"      </center>" & vbNewLine
	end if
end if
Response.Write	"      <form action=""admin_login_short.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
		"      <input type=""hidden"" value=""" & request("target") & """ name=""target"">" & vbNewLine & _
		"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" cellspacing=""1"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""center"" colspan=""2"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Admin Login</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""right"" bgcolor=""" & strPopupTableColor & """ nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;UserName:&nbsp;</font></b></td>" & vbNewLine & _
		"                <td bgcolor=""" & strPopupTableColor & """><input type=""text"" name=""Name"" style=""width:150px;""></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""right"" bgcolor=""" & strPopupTableColor & """ nowrap><b><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:&nbsp;</font></b></td>" & vbNewLine & _
		"                <td bgcolor=""" & strPopupTableColor & """><input type=""Password"" name=""Password"" style=""width:150px;""></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td colspan=""2"" bgcolor=""" & strPopupTableColor & """ align=""center""><input type=""submit"" value=""Login"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      </form>" & vbNewLine
WriteFooterShort
Response.End
%>