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
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Server&nbsp;Date/Time&nbsp;Configuration<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""

	if Err_Msg = "" then
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
	Response.Write	"      <form action=""admin_config_datetime.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
			"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Server Date/Time Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Time Display:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                24hr <input type=""radio"" class=""radio"" name=""strTimeType"" value=""24""" & chkRadio(strTimeType,24,true) & ">" & vbNewLine & _
			"                12hr <input type=""radio"" class=""radio"" name=""strTimeType"" value=""12""" & chkRadio(strTimeType,12,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#timetype')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Time Adjustment:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <select name=""strTimeAdjust"">" & vbNewLine
	for iTimeAdjust = -24 to 24
		Response.Write	"                	<option value=""" & iTimeAdjust & """" & chkSelect(strTimeAdjust,iTimeAdjust) & ">" & iTimeAdjust & "</option>" & vbNewLine
	next
	Response.Write	"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#TimeAdjust')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Current Forum Date/Time:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>&nbsp;" & ChkDate(datetostr(strForumTimeAdjust),"",true) & "</font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Date Display:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
			"                <select name=""strDateType"">" & vbNewLine & _
			"                	<option value=""mdy""" & chkSelect(strDateType,"mdy") & ">12/31/2000 (US short)</option>" & vbNewLine & _
			"                	<option value=""dmy""" & chkSelect(strDateType,"dmy") & ">31/12/2000 (UK short)</option>" & vbNewLine & _
			"                	<option value=""ymd""" & chkSelect(strDateType,"ymd") & ">2000/12/31 (Other short)</option>" & vbNewLine & _
			"                	<option value=""ydm""" & chkSelect(strDateType,"ydm") & ">2000/31/12 (Other short)</option>" & vbNewLine & _
			"                	<option value=""mmdy""" & chkSelect(strDateType,"mmdy") & ">Dec 31 2000 (US med)</option>" & vbNewLine & _
			"                	<option value=""dmmy""" & chkSelect(strDateType,"dmmy") & ">31 Dec 2000 (UK med)</option>" & vbNewLine & _
			"                	<option value=""ymmd""" & chkSelect(strDateType,"ymmd") & ">2000 Dec 31 (Other med)</option>" & vbNewLine & _
			"                	<option value=""ydmm""" & chkSelect(strDateType,"ydmm") & ">2000 31 Dec (Other med)</option>" & vbNewLine & _
			"                	<option value=""mmmdy""" & chkSelect(strDateType,"mmmdy") & ">December 31 2000 (US long)</option>" & vbNewLine & _
			"                	<option value=""dmmmy""" & chkSelect(strDateType,"dmmmy") & ">31 December 2000 (UK long)</option>" & vbNewLine & _
			"                	<option value=""ymmmd""" & chkSelect(strDateType,"ymmmd") & ">2000 December 31 (Other long)</option>" & vbNewLine & _
			"                	<option value=""ydmmm""" & chkSelect(strDateType,"ydmmm") & ">2000 31 December (Other long)</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#datetype')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    </form>" & vbNewLine
end if 
WriteFooter
Response.End
%>
