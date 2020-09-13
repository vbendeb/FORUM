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
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" width=""100%"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;E-mail&nbsp;Server&nbsp;Configuration<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Request.Form("strMailServer") = "" and Request.Form("strMailMode") <> "cdonts" and Request.Form("strEmail") = "1" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the Address of your Mail Server</li>"
	end if
	if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://")) and Request.Form("strEmail") = "1" then
		Err_Msg = Err_Msg & "<li>Do not prefix the Mail Server Address with <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
	end if
	if Request.Form("strSender") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter the E-mail Address of the Forum Administrator</li>"
	else
		if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then 
			Err_Msg = Err_Msg & "<li>You Must enter a valid E-mail Address for the Forum Administrator</li>"
		end if
	end if
	if Request.Form("strRestrictReg") = 1 and Request.Form("strEmailVal") = 0 then
		Err_Msg = Err_Msg & "<li>Email Validation must be enabled in order to enable the Restrict Registration Option</li>"
	end if

	if Err_Msg = "" then
		'## Forum_SQL
		for each key in Request.Form 
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLstring"))
			end if
		next
		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Configuration Posted!</font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""admin_home.asp"">Back To Admin Home</font></a></p>" & vbNewLine
	else
		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
				"      <table align=""center"" border=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
	end if
else
	Dim theComponent(18)
	Dim theComponentName(18)
	Dim theComponentValue(18)

	'## the components
	theComponent(0) = "ABMailer.Mailman"
	theComponent(1) = "Persits.MailSender"
	theComponent(2) = "SMTPsvg.Mailer"
	theComponent(3) = "SMTPsvg.Mailer"
	theComponent(4) = "CDONTS.NewMail"
	theComponent(5) = "CDONTS.NewMail"
	theComponent(6) = "CDO.Message"
	theComponent(7) = "dkQmail.Qmail"
	theComponent(8) = "Dundas.Mailer"
	theComponent(9) = "Dundas.Mailer"
	theComponent(10) = "Geocel.Mailer"
	theComponent(11) = "iismail.iismail.1"
	theComponent(12) = "Jmail.smtpmail"
	theComponent(13) = "MDUserCom.MDUser"
	theComponent(14) = "ASPMail.ASPMailCtrl.1"
	theComponent(15) = "ocxQmail.ocxQmailCtrl.1"
	theComponent(16) = "SoftArtisans.SMTPMail"
	theComponent(17) = "SmtpMail.SmtpMail.1"
	theComponent(18) = "VSEmail.SMTPSendMail"

	'## the name of the components
	theComponentName(0) = "ABMailer v2.2+"
	theComponentName(1) = "ASPEMail"
	theComponentName(2) = "ASPMail"
	theComponentName(3) = "ASPQMail"
	theComponentName(4) = "CDONTS (IIS 3/4/5)"
	theComponentName(5) = "Chili!Mail (Chili!Soft ASP)"
	theComponentName(6) = "CDOSYS (IIS 5/5.1/6)"
	theComponentName(7) = "dkQMail"
	theComponentName(8) = "Dundas Mail (QuickSend)"
	theComponentName(9) = "Dundas Mail (SendMail)"
	theComponentName(10) = "GeoCel"
	theComponentName(11) = "IISMail"
	theComponentName(12) = "JMail"
	theComponentName(13) = "MDaemon"
	theComponentName(14) = "OCXMail"
	theComponentName(15) = "OCXQMail"
	theComponentName(16) = "SA-Smtp Mail"
	theComponentName(17) = "SMTP"
	theComponentName(18) = "VSEmail"

	'## the value of the components
	theComponentValue(0) = "abmailer"
	theComponentValue(1) = "aspemail"
	theComponentValue(2) = "aspmail"
	theComponentValue(3) = "aspqmail"
	theComponentValue(4) = "cdonts"
	theComponentValue(5) = "chilicdonts"
	theComponentValue(6) = "cdosys"
	theComponentValue(7) = "dkqmail"
	theComponentValue(8) = "dundasmailq"
	theComponentValue(9) = "dundasmails"
	theComponentValue(10) = "geocel"
	theComponentValue(11) = "iismail"
	theComponentValue(12) = "jmail"
	theComponentValue(13) = "mdaemon"
	theComponentValue(14) = "ocxmail"
	theComponentValue(15) = "ocxqmail"
	theComponentValue(16) = "sasmtpmail"
	theComponentValue(17) = "smtp"
	theComponentValue(18) = "vsemail"

	Response.Write	"      <form action=""admin_config_email.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
			"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>E-mail Server Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Select E-mail Component:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <select name=""strMailMode"">" & vbNewLine
	dim i, j
	j = 0
	for i=0 to UBound(theComponent)
		if IsObjInstalled(theComponent(i)) then 
			Response.Write	"    <option value=""" & theComponentValue(i) & """" & chkSelect(strMailMode,theComponentValue(i)) & ">" & theComponentName(i) & "</option>" & vbNewline
		else
			j = j + 1
		end if
	next
	if j > UBound(theComponent) then
		Response.Write	"    <option value=""None"">No Compatible Component Found</option>" & vbNewline
	end if 

	Response.Write	"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>E-mail Mode:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strEmail"" value=""1"""
	if j > UBound(theComponent) then Response.Write(" disabled") else if lcase(strEmail) <> "0" then Response.Write(" checked")
	Response.Write	">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strEmail"" value=""0"""
	if j > UBound(theComponent) then Response.Write(" checked") else if lcase(strEmail) = "0" then Response.Write(" checked")
	Response.Write	">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>E-mail Server Address:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <input type=""text"" name=""strMailServer"" size=""25"" value=""" & strMailServer & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#mailserver')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Administrator E-mail Address:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <input type=""text"" name=""strSender"" size=""25"" value=""" & strSender & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#sender')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Require Unique E-mail:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""1""" & chkRadio(strUniqueEmail,1,true) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""0""" & chkRadio(strUniqueEmail,1,false) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#UniqueEmail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>E-mail Validation:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strEmailVal"" value=""1""" & chkRadio(strEmailVal,1,true) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strEmailVal"" value=""0""" & chkRadio(strEmailVal,1,false) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#EmailVal')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Restrict Registration:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""1""" & chkRadio(strRestrictReg,1,true) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""0""" & chkRadio(strRestrictReg,1,false) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#RestrictReg')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Require Logon for sending Mail:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On: <input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""1""" & chkRadio(strLogonForMail,1,true) & ">&nbsp;" & vbNewLine & _
			"                Off: <input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""0""" & chkRadio(strLogonForMail,1,false) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#LogonForMail')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""top"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form>" & vbNewLine
end if 
WriteFooter
Response.End

function IsObjInstalled(strClassString)
	on error resume next
	'## initialize default values
	IsObjInstalled = false
	Err = 0
	'## testing code
	dim xTestObj
	set xTestObj = Server.CreateObject(strClassString)
	if 0 = Err then
		IsObjInstalled = true
	end if
	'## cleanup
	set xTestObj = nothing
	Err = 0
	on error goto 0
end function
%>
