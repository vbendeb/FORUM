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
<!--#INCLUDE FILE="inc_moderation.asp" -->
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
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Feature&nbsp;Configuration<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if Request.Form("Method_Type") = "Write_Configuration" then 
	Err_Msg = ""
	if Request.Form("strIMGInPosts") = "1" and Request.Form("strAllowForumCode") = "0" then 
		Err_Msg = Err_Msg & "<li>Forum Code Must be Enabled in order to Enable Images</li>"
	end if
	if Request.Form("strAllowHTML") = "1" and Request.Form("strAllowForumCode") = "1" then 
		Err_Msg = Err_Msg & "<li>HTML and ForumCode cannot both be On at the same time</li>"
	end if
	if Request.Form("intHotTopicNum") = "" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a Hot Topic Number</li>"
	elseif IsNumeric(Request.Form("intHotTopicNum")) = False then
		Err_Msg = Err_Msg & "<li>Hot Topic Number must be a number</li>"
	elseif cLng(Request.Form("intHotTopicNum")) = 0 then
		Err_Msg = Err_Msg & "<li>Hot Topic Number cannot be 0</li>"
	end if
	if left(Request.Form("intHotTopicNum"), 1) = "-" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number</li>"
	end if
	if left(Request.Form("intHotTopicNum"), 1) = "+" then 
		Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number without the <b>+</b></li>"
	end if
	if Request.Form("strPageSize") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter the number of Items per Page</li>"
	elseif IsNumeric(Request.Form("strPageSize")) = False then
		Err_Msg = Err_Msg & "<li>Items per Page must be a number</li>"
	elseif cLng(Request.Form("strPageSize")) = 0 then
		Err_Msg = Err_Msg & "<li>Items per Page cannot be 0</li>"
	end if
	if Request.Form("strPageNumberSize") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter the number of Pages per Row</li>"
	elseif IsNumeric(Request.Form("strPageNumberSize")) = False then
		Err_Msg = Err_Msg & "<li>Pages per Row must be a number</li>"
	elseif cLng(Request.Form("strPageNumberSize")) = 0 then
		Err_Msg = Err_Msg & "<li>Pages per Row cannot be 0</li>"
	end if

	if (strShowTimer = "1" or Request.Form("strShowTimer") = "1") and Request.Form("strShowTimer") <> "0" then
		if trim(Request.Form("strTimerPhrase")) = "" then
			Err_Msg = Err_Msg & "<li>You Must Enter a Phrase for the Timer</li>"
		end if
		if Instr(Request.Form("strTimerPhrase"), "[TIMER]") = "0" then
			Err_Msg = Err_Msg & "<li>Your Timer Phrase must contain the [TIMER] placeholder</li>"
		end if
	end if
	if strModeration = "1" and Request.Form("strModeration") = "0" then
        	if CheckForUnmoderatedPosts("BOARD", 0, 0, 0) > 0 then
			Err_Msg = Err_Msg & "<li>Please Approve or Delete all UnModerated/Held posts before turning Moderation off.</li>"
		end if
	end if

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
	Response.Write	"      <form action=""admin_config_features.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewLine & _
			"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
			"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strHeadCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>Feature Configuration</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Security Settings</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Secure Admin Mode:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strSecureAdmin"" value=""1""" & chkRadio(strSecureAdmin,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strSecureAdmin"" value=""0""" & chkRadio(strSecureAdmin,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#secureadminmode')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Non-Cookie Mode:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strNoCookies"" value=""1""" & chkRadio(strNoCookies,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strNoCookies"" value=""0""" & chkRadio(strNoCookies,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#allownoncookies')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>General Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>IP Logging:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strIPLogging"" value=""1""" & chkRadio(strIPLogging,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strIPLogging"" value=""0""" & chkRadio(strIPLogging,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#IPLogging')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Flood Control:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strFloodCheck"" value=""1""" & chkRadio(strFloodCheck,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strFloodCheck"" value=""0""" & chkRadio(strFloodCheck,0,true) & ">" & vbNewLine & _
			"                <select name=""strFloodCheckTime"">" & vbNewLine & _
			"                	<option value=""-30""" & chkSelect(strFloodCheckTime,-30) & ">30 seconds</option>" & vbNewLine & _
			"                	<option value=""-60""" & chkSelect(strFloodCheckTime,-60) & ">60 seconds</option>" & vbNewLine & _
			"                	<option value=""-90""" & chkSelect(strFloodCheckTime,-90) & ">90 seconds</option>" & vbNewLine & _
			"                	<option value=""-120""" & chkSelect(strFloodCheckTime,-120) & ">120 seconds</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#FloodCheck')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a>&nbsp;</td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Private Forums:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strPrivateForums"" value=""1""" & chkRadio(strPrivateForums,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strPrivateForums"" value=""0""" & chkRadio(strPrivateForums,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#privateforums')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Group Categories:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strGroupCategories"" value=""1""" & chkRadio(strGroupCategories,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strGroupCategories"" value=""0""" & chkRadio(strGroupCategories,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#groupcategories')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Highest level of Subscription:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <select name=""strSubscription"">" & vbNewLine & _
			"                	<option value=""0""" & chkSelect(strSubscription,0) & ">No Subscriptions Allowed</option>" & vbNewLine & _
			"                	<option value=""1""" & chkSelect(strSubscription,1) & ">Subscribe to Whole Board</option>" & vbNewLine & _
			"                	<option value=""2""" & chkSelect(strSubscription,2) & ">Subscribe by Category</option>" & vbNewLine & _
			"                	<option value=""3""" & chkSelect(strSubscription,3) & ">Subscribe by Forum</option>" & vbNewLine & _
			"                	<option value=""4""" & chkSelect(strSubscription,4) & ">Subscribe by Topic</option>" & vbNewLine & _
			"                </select>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#Subscription')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Bad Word Filter:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strBadWordFilter"" value=""1""" & chkRadio(strBadWordFilter,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strBadWordFilter"" value=""0""" & chkRadio(strBadWordFilter,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#badwordfilter')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Moderation Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow Topic Moderation:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strModeration"" value=""1""" & chkRadio(strModeration,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strModeration"" value=""0""" & chkRadio(strModeration,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#Moderation')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Moderators:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowModerators"" value=""1""" & chkRadio(strShowModerators,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowModerators"" value=""0""" & chkRadio(strShowModerators,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowModerator')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Restrict Moderators to&nbsp;&nbsp;<br /> moving their own topics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strMoveTopicMode"" value=""1""" & chkRadio(strMoveTopicMode,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strMoveTopicMode"" value=""0""" & chkRadio(strMoveTopicMode,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#MoveTopicMode')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>AutoEmail author&nbsp;&nbsp;<br />when moving topics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strMoveNotify"" value=""1""" & chkRadio(strMoveNotify,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strMoveNotify"" value=""0""" & chkRadio(strMoveNotify,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#MoveNotify')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Forum Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Archive Functions:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strArchiveState"" value=""1""" & chkRadio(strArchiveState,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strArchiveState"" value=""0""" & chkRadio(strArchiveState,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ArchiveState')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Detailed Statistics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowStatistics"" value=""1""" & chkRadio(strShowStatistics,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowStatistics"" value=""0""" & chkRadio(strShowStatistics,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#stats')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Jump To Last Post Link:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strJumpLastPost"" value=""1""" & chkRadio(strJumpLastPost,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strJumpLastPost"" value=""0""" & chkRadio(strJumpLastPost,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#JumpLastPost')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Quick Paging:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowPaging"" value=""1""" & chkRadio(strShowPaging,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowPaging"" value=""0""" & chkRadio(strShowPaging,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowPaging')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Pagenumbers per row:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <input type=""text"" name=""strPageNumberSize"" size=""5"" maxLength=""3"" value=""" & chkExistElse(strPageNumbersize,10) & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#pagenumbersize')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Topic Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow Sticky Topics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strStickyTopic"" value=""1""" & chkRadio(strStickyTopic,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strStickyTopic"" value=""0""" & chkRadio(strStickyTopic,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#StickyTopic')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Edited By on Date:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strEditedByDate"" value=""1""" & chkRadio(strEditedByDate,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strEditedByDate"" value=""0""" & chkRadio(strEditedByDate,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#editedbydate')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Prev / Next Topic:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowTopicNav"" value=""1""" & chkRadio(strShowTopicNav,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowTopicNav"" value=""0""" & chkRadio(strShowTopicNav,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowTopicNav')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Send Topic to a Friend Link:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowSendToFriend"" value=""1""" & chkRadio(strShowSendToFriend,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowSendToFriend"" value=""0""" & chkRadio(strShowSendToFriend,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowSendToFriend')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Printer Friendly Link:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowPrinterFriendly"" value=""1""" & chkRadio(strShowPrinterFriendly,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowPrinterFriendly"" value=""0""" & chkRadio(strShowPrinterFriendly,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowPrinterFriendly')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Hot Topics:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strHotTopic"" value=""1""" & chkRadio(strHotTopic,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strHotTopic"" value=""0""" & chkRadio(strHotTopic,0,true) & ">" & vbNewLine & _
			"                <input type=""text"" name=""intHotTopicNum"" size=""5"" maxLength=""3"" value=""" & chkExistElse(intHotTopicNum,20) & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#hottopics')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Items per page:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <input type=""text"" name=""strPageSize"" size=""5"" maxLength=""3"" value=""" & chkExistElse(strPageSize,15) & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#pagesize')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Posting Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow HTML:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strAllowHTML"" value=""1""" & chkRadio(strAllowHTML,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strAllowHTML"" value=""0""" & chkRadio(strAllowHTML,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#AllowHTML')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow Forum Code:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strAllowForumCode"" value=""1""" & chkRadio(strAllowForumCode,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strAllowForumCode"" value=""0""" & chkRadio(strAllowForumCode,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#AllowForumCode')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Images in Posts:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strIMGInPosts"" value=""1""" & chkRadio(strIMGInPosts,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strIMGInPosts"" value=""0""" & chkRadio(strIMGInPosts,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#imginposts')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Icons:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strIcons"" value=""1""" & chkRadio(strIcons,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strIcons"" value=""0""" & chkRadio(strIcons,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#icons')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow Signatures:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strSignatures"" value=""1""" & chkRadio(strSignatures,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strSignatures"" value=""0""" & chkRadio(strSignatures,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#signatures')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Allow Dynamic Signatures:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strDSignatures"" value=""1""" & chkRadio(strDSignatures,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strDSignatures"" value=""0""" & chkRadio(strDSignatures,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#dsignatures')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Format Buttons:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowFormatButtons"" value=""1""" & chkRadio(strShowFormatButtons,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowFormatButtons"" value=""0""" & chkRadio(strShowFormatButtons,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowFormatButtons')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Smilies Table:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowSmiliesTable"" value=""1""" & chkRadio(strShowSmiliesTable,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowSmiliesTable"" value=""0""" & chkRadio(strShowSmiliesTable,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowSmiliesTable')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Quick Reply:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowQuickReply"" value=""1""" & chkRadio(strShowQuickReply,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowQuickReply"" value=""0""" & chkRadio(strShowQuickReply,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowQuickReply')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2"" align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Misc Features</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Show Timer:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                On:<input type=""radio"" class=""radio"" name=""strShowTimer"" value=""1""" & chkRadio(strShowTimer,0,false) & ">&nbsp;" & vbNewLine & _
			"                Off:<input type=""radio"" class=""radio"" name=""strShowTimer"" value=""0""" & chkRadio(strShowTimer,0,true) & ">" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#timer')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Timer Phrase:</b>&nbsp;</font></td>" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <input type=""text"" name=""strTimerPhrase"" size=""45"" maxLength=""50"" value=""" & chkExistElse(strTimerPhrase,"This page was generated in [TIMER] seconds.") & """>" & vbNewLine & _
			"                <a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#timerphrase')"">" & getCurrentIcon(strIconSmileQuestion,"","") & "</a></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr valign=""middle"">" & vbNewLine & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      </form>" & vbNewLine
end if 
WriteFooter
Response.end
%>
