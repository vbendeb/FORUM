<%@CODEPAGE=1251%>
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
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<%
Response.Write	"      <table width=""100%"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">���� �����</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;������ ��� ������?<br />" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if lcase(strEmail) <> "1" then
	Response.Redirect("default.asp")
end if

if Request.Form("mode") <> "DoIt" and Request.Form("mode") <> "UpdateIt" and Request.QueryString("pwkey") = "" then
	call ShowForm
elseif Request.QueryString("pwkey") <> "" and Request.Form("mode") <> "UpdateID" then
	key = chkString(Request.QueryString("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>������������ ������!</b></font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>������, ������� �� ����� ���������� �� ���, ��� �������� � ����� ���� ������.<br />���������� ������ ���� ��� ������������ � ����������� ����� ��� ���, ����� ������� �� ������ ""����� ������P?"" �� ������� �������� ������.<br />���� �������� �� ������� - ��������� � <a href=""mailto:" & strSender & """>���������������</a> ������.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""default.asp"">����� �� �����</font></a></p>" & vbNewLine
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>������������ ������!</b></font></p>" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>������, ������� �� ����� ���������� �� ���, ��� �������� � ����� ���� ������.<br />���������� ������ ���� ��� ������������ � ����������� ����� ��� ���, ����� ������� �� ������ ""����� ������P?"" �� ������� �������� ������.<br />���� �������� �� ������� - ��������� � <a href=""mailto:" & strSender & """>���������������</a>  ������.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""default.asp"">����� �� �����</font></a></p>" & vbNewLine
	else
		PWMember_ID = rsKey("MEMBER_ID")
		call showForm2
	end if

	rsKey.close
	set rsKey = nothing
elseif Request.Form("pwkey") <> "" and Request.Form("mode") = "UpdateIt" then
	key = chkString(Request.Form("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>������������ ������!</b></font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "������, ������� �� ����� ���������� �� ���, ��� �������� � ����� ���� ������..<br />���������� ������ ���� ��� ������������ � ����������� ����� ��� ���, ����� ������� �� ������ ""����� ������P?"" �� ������� �������� ������.<br />���� �������� �� ������� - ��������� � <a href=""mailto:" & strSender & """>���������������</a> ������.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""default.asp"">" & strBackToForum & "</font></a></p>" & vbNewLine
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>Your password key did not match!</b></font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "������, ������� �� ����� ���������� �� ���, ��� �������� � ����� ���� ������.<br />���������� ������ ���� ��� ������������ � ����������� ����� ��� ���, ����� ������� �� ������ ""����� ������P?"" �� ������� �������� ������.<br />���� �������� �� ������� - ��������� � <a href=""mailto:" & strSender & """>���������������</a> ������.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""default.asp"">" & strBackToForum & "</font></a></p>" & vbNewLine
        else
		if trim(Request.Form("Password")) = "" then
			Err_Msg = Err_Msg & "<li>�� ������� ������������ �������</li>"
		end if
		if Len(Request.Form("Password")) > 25 then
			Err_Msg = Err_Msg & "<li>����� ������ �� ����� ��������� 25 ��������</li>"
		end if
		if Request.Form("Password") <> Request.Form("Password2") then
			Err_Msg = Err_Msg & "<li>���� ������ ������������.</li>"
		end if

		if Err_Msg = "" then
			strEncodedPassword = sha256("" & Request.Form("Password"))
			pwkey = ""

			'Update the user's password
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " SET M_PASSWORD = '" & chkString(strEncodedPassword,"SQLString") & "'"
			strSql = strSql & ", M_PWKEY = '" & chkString(pwkey,"SQLString") & "'"
			strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		else
			if Err_Msg <> "" then 
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>�������� � ������ �������</font></p>" & vbNewLine & _
						"      <table align=""center"" border=""0"">" & vbNewLine & _
						"        <tr>" & vbNewLine & _
						"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
						"        </tr>" & vbNewLine & _
						"      </table>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
				rsKey.close
				set rsKey = nothing
				WriteFooter
				Response.End 
			end if
		end if
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Your Password has been updated!</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "You may now login"
		if strAuthType = "db" then Response.Write(" ��������� ����� ������ ������������ � ����� �������")
		Response.Write	".</font></p>" & vbNewLine
		Response.Write	"      <meta http-equiv=""Refresh"" content=""2; URL=default.asp"">" & vbNewLine
		Response.Write	"      " & strParagraphFormat1 & "<a href=""default.asp"">" & strBackToForum & "</font></a></p>" & vbNewLine
	end if

	rsKey.close
	set rsKey = nothing
else
	Err_Msg = ""

	if trim(Request.Form("Name")) = "" then
		Err_Msg = Err_Msg & "<li>��� ������������ �����������</li>"
	end if

	if trim(Request.Form("Email")) = "" then
		Err_Msg = Err_Msg & "<li>����� ����������� ����� ����������</li>"
	end if

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & ChkString(Trim(Request.Form("Name")), "SQLString") &"'"
	strSql = strSql & " AND M_EMAIL = '" & ChkString(Trim(Request.Form("Email")), "SQLString") &"'"

	set rs = my_Conn.Execute (strSql)

	if rs.BOF and rs.EOF then
		Err_Msg = Err_Msg & "<li>���� ��� ������������, ���� ����������� ����� ����������� � ����� ���� ������.</li>"
	else
		PWMember_ID = rs("MEMBER_ID")
		PWMember_Name = rs("M_NAME")
		PWMember_Email = rs("M_EMAIL")
	end if
	
	rs.close
	set rs = nothing

	if Err_Msg = "" then
		pwkey = GetKey("none")

		'Update the user Member Level
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_PWKEY = '" & chkString(pwkey,"SQLString") & "'"
		strSql = strSql & " WHERE MEMBER_ID = " & PWMember_ID

		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

		if lcase(strEmail) = "1" then
			'## E-mails Message to the Author of this Reply.  
			strRecipientsName = PWMember_Name
			strRecipients = PWMember_Email
			strFrom = strSender
			strFromName = strForumTitle
			strsubject = strForumTitle & " - ������ ������?? "
			strMessage = "Hello " & PWMember_Name & vbNewline & vbNewline
			strMessage = strMessage & "��� ��������� ������� ��� �� " & strForumTitle & " ������ ��� �� ��������� ������ ����� �������� ""������ ������?""." & vbNewline & vbNewline
			strMessage = strMessage & "���������� ������� �� ������ �����, ����� ��������� �������." & vbNewline & vbNewLine
			strMessage = strMessage & strForumURL & "password.asp?pwkey=" & pwkey & vbNewline & vbNewline
			strMessage = strMessage & vbNewLine & "���� �� �� ������ ���� ������ � �������� ��� ��������� �� ������ - ������� �������������� �������� � ����� ������� �� ���������." & vbNewLine & vbNewLine
%>
			<!--#INCLUDE FILE="inc_mail.asp" -->
<%
		end if
	else
		if Err_Msg <> "" then 
			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>�������� � ����� �����������</font></p>" & vbNewLine & _
					"      <table align=""center"" border=""0"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">��������� � ����� ������</a></font></p>" & vbNewLine
			WriteFooter
			Response.End 
		end if
	end if
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>������ ����� ��������!</font></p>" & vbNewLine & _
			"      " & strParagraphFormat1 & "���������� �������� �������� ������������� � ��������� ��������� �� ������ <b>" & ChkString(PWMember_Email,"") & "</b> ����� ��������� ���� �������.</font></p>" & vbNewLine
	Response.Write	"      <meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & vbNewLine
	Response.Write	"      " & strParagraphFormat1 & "<a href=""default.asp"">����� �� �����</font></a></p>" & vbNewLine
end if 
WriteFooter
Response.End

sub ShowForm()
	Response.Write	"      <form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input name=""mode"" type=""hidden"" value=""DoIt"">" & vbNewLine & _
			"      <table width=""100%"" border=""0"" align=""center"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewline & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewline & _
			"              <tr>" & vbNewline & _
			"                <td colspan=""2"" align=""center"" bgcolor=""" & strHeadCellColor & """ valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>������ ������?</font></b></td>" & vbNewline & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewline & _
			"                <td colspan=""2"" align=""left"" bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>���� ������� �������� � ���� 3 ����:" & vbNewLine & _
			"                <ul>" & vbNewLine & _
			"                 <font color=""" & strHiLiteFontColor & """><li><b>������ ���:</b><br />������� ���� ����� ���/��� � e-mail ����� � ����� ��� ��������� e-mail ��������� � ����� ��� ����� ������������� � ������� ������ Submit.</li></font>" & vbNewLine & _
			"                <li><b>������ ���:</b><br />��������� ���� e-mail ����� � ������� �� ���� � ��������� ��� �������� �� ��� ��������.</li>" & vbNewLine & _
			"                <li><b>������ ���:</b><br />�������� ����� ������.</li>" & vbNewLine & _
			"                </ul></font></td>" & vbNewline & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td width=""50%"" align=""right"" bgcolor=""" & strForumCellColor & """ nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;UserName:&nbsp;</font></b></td>" & vbNewLine & _
			"                <td width=""50%"" bgcolor=""" & strForumCellColor & """><input type=""text"" name=""Name"" size=""25"" maxLength=""25""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td width=""50%"" align=""right"" bgcolor=""" & strForumCellColor & """ nowrap><b><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>E-mail Address:&nbsp;</font></b></td>" & vbNewLine & _
			"                <td width=""50%"" bgcolor=""" & strForumCellColor & """><input type=""text"" name=""Email"" size=""25"" maxLength=""50""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td colspan=""2"" bgcolor=""" & strForumCellColor & """ align=""center""><input type=""submit"" value=""Submit"" id=""Submit1"" name=""Submit1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Reset"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form><br />" & vbNewLine
end sub

sub ShowForm2()
	Response.Write	"      <form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input name=""mode"" type=""hidden"" value=""UpdateIt"">" & vbNewLine & _
			"      <input name=""MEMBER_ID"" type=""hidden"" value=""" & PWMember_ID & """>" & vbNewLine & _
			"      <input name=""pwkey"" type=""hidden"" value=""" & key & """>" & vbNewLine & _
			"      <table width=""100%"" border=""0"" align=""center"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewline & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewline & _
			"              <tr>" & vbNewline & _
			"                <td colspan=""2"" align=""center"" bgcolor=""" & strHeadCellColor & """ valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Forgot your Password?</font></b></td>" & vbNewline & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewline & _
			"                <td colspan=""2"" align=""left"" bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>This is a 3 step process:" & vbNewLine & _
			"                <ul>" & vbNewLine & _
			"                <li><b>First Step:</b><br />Enter your username and e-mail address in the form below to receive an e-mail containing a code to verify that you are who you say you are. <b>(COMPLETED)</b></li>" & vbNewLine & _
			"                <li><b>Second Step:</b><br />Check your e-mail and then click on the link that is provided to return to this page. <b>(COMPLETED)</b></li>" & vbNewLine & _
			"                <font color=""" & strHiLiteFontColor & """><li><b>Third Step:</b><br />Choose your new password.</li></font>" & vbNewLine & _
			"                </ul></font></td>" & vbNewline & _
			"              </tr>" & vbNewLine & _
       			"              <tr>" & vbNewLine & _
			"                <td width=""50%"" bgColor=""" & strForumCellColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:&nbsp;</font></b></td>" & vbNewLine & _
			"                <td width=""50%"" bgColor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><input name=""Password"" type=""Password"" size=""25"" maxLength=""25"" value=""""></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
       			"              <tr>" & vbNewLine & _
			"                <td width=""50%"" bgColor=""" & strForumCellColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password Again:&nbsp;</font></b></td>" & vbNewLine & _
			"                <td width=""50%"" bgColor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><input name=""Password2"" type=""Password"" maxLength=""25"" size=""25"" value=""""></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td colspan=""2"" bgcolor=""" & strForumCellColor & """ align=""center""><input type=""submit"" value=""Submit"" id=""Submit1"" name=""Submit1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Reset"" id=""Submit1"" name=""Submit1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form><br />" & vbNewLine
end sub
%>