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
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<!--#INCLUDE file="inc_func_member.asp" -->
<%
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""
	if (Request.Form("YName") = "") then 
		Err_Msg = Err_Msg & "<li>You must enter your name!</li>"
	end if
	if (Request.Form("YEmail") = "") then 
		Err_Msg = Err_Msg & "<li>You Must give your e-mail address</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address</li>"
		end if
	end if
	if (Request.Form("Name") = "") then 
		Err_Msg = Err_Msg & "<li>You must enter the recipients name</li>"
	end if
	if (Request.Form("Email") = "") then 
		Err_Msg = Err_Msg & "<li>You Must enter the recipients e-mail address</li>"
	else
		if (EmailField(Request.Form("Email")) = 0) then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address for the recipient</li>"
		end if
	end if
	if (Request.Form("Msg") = "") then 
		Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
	end if
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strSubject = "From: " & Request.Form("YName") & " Interesting Page"
			strMessage = "Hello " & Request.Form("Name") & vbNewline & vbNewline
			strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline
			strMessage = strMessage & "You received this from : " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") "

			if Request.Form("YEmail") <> "" then 
				strSender = Request.Form("YEmail")
			end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			Response.Write("      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>E-mail has been sent</font></p>" & vbNewLine)
		else
			Response.Write("      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your E-mail</font></p>" & vbNewLine)
			Response.Write	"      <table>" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine
			Response.Write("      <p><font size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine)
		end if
	end if
else 
	'## Forum_SQL
	strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & chkString(strDBNTUserName,"SQLString") & "'"

	set rs = my_conn.Execute (strSql)
	YName = ""
	YEmail = ""

	if (rs.EOF or rs.BOF)  then
		if strLogonForMail <> "0" then
			Err_Msg = Err_Msg & "<li>You Must be logged on to send a message</li>"
			Response.Write("      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your E-mail</font></p>" & vbNewLine)
			Response.Write	"      <table>" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine
			set rs = nothing
			Response.Write("      <p><font size=""" & strDefaultFontSize & """><a href=""JavaScript:onClick= window.close()"">Close Window</A></font></p>" & vbNewLine)
			Response.End
		end if
	else
	  	YName = Trim("" & rs("M_NAME"))
	  	YEmail = Trim("" & rs("M_EMAIL"))
	end if

	rs.close
	set rs = nothing

	Response.Write("      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Send Topic to a Friend</font></p>" & vbNewLine)

	Response.Write	"      <form action=""pop_send_to_friend.asp?mode=DoIt"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input type=""hidden"" name=""Page"" value=""" & Request.QueryString & """>" & vbNewLine & _
			"      <table border=""0"" width=""90%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send To Name:</font></b></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""Name"" size=""25""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send To E-mail:</font></b></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""Email"" size=""25""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Name:</font></b></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input name=""YName"" type="""
			if YName <> "" then
				Response.Write("hidden")
			else
				Response.Write("text")
			end if
	Response.Write	""" value=""" & YName & """ size=""25""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
			if YName <> "" then
				Response.Write(YName)
			end if
	Response.Write	"</font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your E-mail:</font></b></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><input name=""YEmail"" type="""
			if YEmail <> "" then
				Response.Write("hidden")
			else
				Response.Write("text")
			end if
	Response.Write	""" value=""" & YEmail & """ size=""25""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
			if YEmail <> "" then
				Response.Write(YEmail)
			end if
	Response.Write	"</font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Message:</font></b></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><textarea name=""Msg"" cols=""40"" rows=""5"" readonly>I thought you might be interested in this post:" & vbNewline & vbNewline & Request.QueryString("url") & "</textarea></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form>" & vbNewLine
end if
WriteFooterShort
Response.End
%>