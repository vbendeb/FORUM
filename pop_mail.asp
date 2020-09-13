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
if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

'## Forum_SQL
strSql = "SELECT M.M_RECEIVE_EMAIL, M.M_EMAIL, M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

set rs = my_Conn.Execute (strSql)

Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Send an E-MAIL Message</font></p>" & vbNewLine

if rs.bof or rs.eof then
	Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>There is no Member with that Member ID</font></p>" & vbNewLine
else
	strRName = ChkString(rs("M_NAME"),"display")
	if mLev > 2 or rs("M_RECEIVE_EMAIL") = "1" then
		if lcase(strEmail) = "1" then
			if Request.QueryString("mode") = "DoIt" then
				Err_Msg = ""
				if Request.Form("YName") = "" then 
					Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
				end if
				if Request.Form("YEmail") = "" then 
					Err_Msg = Err_Msg & "<li>You Must give your e-mail address</li>"
				else
					if EmailField(Request.Form("YEmail")) = 0 then 
						Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address</li>"
					end if
				end if
				if Request.Form("Name") = "" then 
					Err_Msg = Err_Msg & "<li>You must enter the recipients name</li>"
				end if
				if Request.Form("Msg") = "" then 
					Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
				end if
				'##  E-mails Message to the Author of this Reply.  
				if (Err_Msg = "") then
					strRecipientsName = strRName
					strRecipients = rs("M_EMAIL")
					strFrom = Request.Form("YEmail")
					strFromName = Request.Form("YName")
					strSubject = "Sent From " & strForumTitle & " by " & Request.Form("YName")
					strMessage = "Hello " & strRName & vbNewline & vbNewline
					strMessage = strMessage & "You received the following message from: " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") " & vbNewline & vbNewline 
					strMessage = strMessage & "At: " & strForumURL & vbNewline & vbNewline
					strMessage = strMessage & Request.Form("Msg") & vbNewline & vbNewline

					if strFrom <> "" then 
						strSender = strFrom
					end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>E-mail has been sent</font></p>" & vbNewLine
				else
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your E-mail</font></p>" & vbNewLine
					Response.Write	"      <table>" & vbNewLine & _
							"        <tr>" & vbNewLine & _
							"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
							"        </tr>" & vbNewLine & _
							"      </table>" & vbNewLine & _
							"    <p><font size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
					WriteFooterShort
					Response.End 
				end if
			else 
				Err_Msg = ""
				if rs("M_EMAIL") <> " " then
			     		strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
				     	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
				     	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & chkString(strDBNTUserName,"SQLString") & "'"

					set rs2 = my_conn.Execute (strSql)
					YName = ""
					YEmail = ""

					if (rs2.EOF or rs2.BOF)  then
						if strLogonForMail <> "0" then 
							Err_Msg = Err_Msg & "<li>You must be logged on to send a message</li>"

							Response.Write	"      <table>" & vbNewLine & _
									"        <tr>" & vbNewLine & _
									"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
									"        </tr>" & vbNewLine & _
									"      </table>" & vbNewLine
							WriteFooterShort
							Response.End
						end if
					else
						YName = Trim("" & rs2("M_NAME"))
						YEmail = Trim("" & rs2("M_EMAIL"))
					end if
					rs2.close
					set rs2 = nothing

					Response.Write	"      <form action=""pop_mail.asp?mode=DoIt&id=" & intMemberID & """ method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
							"      <input type=""hidden"" name=""Page"" value=""" & Request.QueryString("page") & """>" & vbNewLine & _
							"      <table border=""0"" width=""90%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
							"        <tr>" & vbNewLine & _
							"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
							"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send To Name:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & strRName & "<input type=""hidden"" name=""Name"" value=""" & strRName & """></font></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Name:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """><input name=""YName"" type="""
					if YName <> "" then Response.Write("hidden") else Response.Write("text")
					Response.Write	""" value=""" & YName & """ size=""25""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if YName <> "" then Response.Write(YName)
					Response.Write	"</font></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your E-mail:</font></b></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """><input name=""YEmail"" type="""
					if YEmail <> "" then Response.Write("hidden") else Response.Write("text")
					Response.Write	""" value=""" & YEmail & """ size=""25""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>"
					if YEmail <> "" then Response.Write(YEmail)
					Response.Write	"</font></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Message:</font></b></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2""><textarea name=""Msg"" cols=""40"" rows=""5""></textarea></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"            </table>" & vbNewLine & _
							"          </td>" & vbNewLine & _
							"        </tr>" & vbNewLine & _
							"      </table>" & vbNewLine & _
							"      </form>" & vbNewLine
				else
					Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>No E-mail address is available for this user.</font></p>" & vbNewLine
				end if
			end if
		else
			Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Click to send <a href=""mailto:" & rs("M_EMAIL") & """>" & strRName & "</a> an e-mail</font></p>" & vbNewLine
		end if
	else
		Response.Write	"    <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>This Member does not wish to receive e-mail.</font></p>" & vbNewLine
	end if
end if
set rs = nothing
WriteFooterShort
Response.End
%>