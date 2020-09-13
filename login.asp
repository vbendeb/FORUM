<%@CODEPAGE=1251%>
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
if MemberID > 0 then Response.Redirect("default.asp")
Response.Write	"      <table border=""0"" width=""100%"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"All Forums","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"Forum Login","") & "&nbsp;Member&nbsp;Login<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

fName = strDBNTFUserName
fPassword = ChkString(Request.Form("Password"), "SQLString")

RequestMethod = Request.ServerVariables("Request_Method")

if RequestMethod = "POST" Then
	strEncodedPassword = sha256("" & fPassword)
	select case chkUser(fName, strEncodedPassword,-1)
		case 1, 2, 3, 4
			Call DoCookies(Request.Form("SavePassword"))
			strLoginStatus = 1
		case else
			strLoginStatus = 0
	end select

	if strLoginStatus = 1 then
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Добро Пожаловать!</font></p>" & vbNewLine
		Response.Write	"      " & strParagraphFormat1 & "<a href="""
		if Request("target") = "" then
			Response.Write	"default.asp"
		else
			Response.Write	request("target")
		end if
		Response.Write	""">Click here to Continue</a></font></p>" & vbNewLine

		Response.Write	"      <meta http-equiv=""Refresh"" content=""2; URL="
		if Request("target") = "" then
			Response.Write	"default.asp"
		else
			Response.Write	request("target")
		end if
		Response.Write	""">" & vbNewline & _
				"      <br />" & vbNewLine

		WriteFooter
		Response.End
	end if
end if
Response.Write	"      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"        <form action=""login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
		"        <input type=""hidden"" value=""" & chkString(request("target"),"display") & """ name=""target"">" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""4"" align=""center"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Member Login</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Member Login</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """>" & vbNewLine & _
		"                  <table border=""0"" cellpadding=""6"" cellspacing=""0"" width=""90%"" align=""center"">" & vbNewLine & _
		"                    <tr valign=""top"">" & vbNewLine & _
		"                      <td width=""49%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine
if RequestMethod = "POST" and strLoginStatus = 0 then Response.Write("                      <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>Такой записи нет. Повторите ввод</font><br />" & vbNewLine) else Response.Write("<br />" & vbNewLine)
Response.Write	"                      <b>Логин / Пароль:</b></font>" & vbNewLine & _
		"                        <table border=""0"" cellpadding=""2"" cellspacing=""0"">" & vbNewLine & _
		"                          <tr>" & vbNewLine & _
		"                            <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                            Логин (имя/ник:)<br />" & vbNewLine & _
		"                            <input type=""text"" name=""Name"" size=""20"" maxLength=""25"" tabindex=""1"" value="""" style=""width:150px;""></td>" & vbNewLine & _
		"                            <td rowspan=""2"" valign=""bottom""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine
if strGfxButtons = "1" then
	Response.Write	"                            <input src=""" & strImageUrl & "button_login.gif"" type=""image"" border=""0"" value=""Login"" id=""submit1"" name=""submit1"" tabindex=""3""></font></td>" & vbNewLine
else
	Response.Write	"                            <input class=""button"" type=""submit"" value=""Login"" id=""submit1"" name=""submit1"" tabindex=""3""></font></td>" & vbNewLine
end if 
Response.Write	"                          </tr>" & vbNewLine & _
		"                          <tr>" & vbNewLine & _
		"                            <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                            Пароль:<br />" & vbNewLine & _
		"                            <input type=""password"" name=""Password"" size=""20"" tabindex=""2"" maxLength=""25"" value="""" style=""width:150px;""></td>" & vbNewLine & _
		"                          </tr>" & vbNewLine & _
		"                          <tr>" & vbNewLine & _
		"                            <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                            <input type=""checkbox"" name=""SavePassWord"" tabindex=""4"" value=""true"" checked> Запомнить пароль</font></td>" & vbNewLine & _
		"                          </tr>" & vbNewLine & _
		"                        </table>" & vbNewLine & _
		"                      </td>" & vbNewLine & _
		"                      <script language=""JavaScript"" type=""text/javascript"">document.Form1.Name.focus();</script>" & vbNewLine & _
		"                      <td width=""2%""nowrap></td>" & vbNewLine & _
		"                      <td width=""49%""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><br /><b>Вопросы:</b><br />" & vbNewLine & _
		"                      <span style=""font-size: 6px;""><br /></span>" & vbNewLine & _
		"                      <acronym title=""Зачем нужна регистрация?""><span class=""spnMessageText""><a href=""faq.asp#register""" & dWStatus("Зачем нужна регистрация?") & ">Зачем нужна регистрация?</a></span></acronym><br />" & vbNewLine
if strEmail = "1" then Response.Write("                      <acronym title=""Choose a new password if you have forgotten your current one.""><span class=""spnMessageText""><a href=""password.asp""" & dWStatus("Choose a new password if you have forgotten your current one.") & ">Забыли Ваш пароль?</a></span></acronym><br /><br />" & vbNewLine) else Response.Write("                      <br />" & vbNewLine)
Response.Write	"                      Я здесь впервые - как зарегистрироваться?<br />"
if strProhibitNewMembers = "1" then
	Response.Write	"<font size=""" & strFooterFontSize & """ color=""" & strHiLiteFontColor & """>The Administrator has turned off Registration for this forum.<br />Only registered members are able to log in</font></font></td>" & vbNewLine
else
	Response.Write	"<acronym title=""Регистрация""><span class=""spnMessageText""><a href=""policy.asp""" & dWStatus("Регистрация") & ">Регистрация нового пользователя</a></span></acronymn></font></td>" & vbNewLine
end if
Response.Write	"                    </tr>" & vbNewLine & _
		"                  </table>" & vbNewLine & _
		"                </td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </form>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine
WriteFooter
%>
