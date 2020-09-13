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
Response.Write 	"          <!-- START JUMP TO -->" & vbNewline
Response.Write	"          <form name=""Stuff"">" & vbNewLine
Response.Write	"          <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Перейти сразу на:</b></font>" & vbNewline
Response.Write	"          <select name=""SelectMenu"" size=""1"" onchange=""if(this.options[this.selectedIndex].value != '' ){ jumpTo(this) }"">" & vbNewline & _
		"          	<option value="""">Выбор Форума</option>" & vbNewline
'## Get all Forum Categories From DB
if IsEmpty(Application(strCookieURL & "JumpBoxChanged")) then
	strJumpBoxChanged = Session(strCookieURL & "JumpBoxDate")
else
	strJumpBoxChanged = Application(strCookieURL & "JumpBoxChanged")
end if
if IsEmpty(Session(strCookieURL & "JumpBox")) or (strJumpBoxChanged > Session(strCookieURL & "JumpBoxDate")) then
	Dim strSelectBox
	if allAllowedForums = "" or isNull(allAllowedForums) then
		if strPrivateForums = "1" and mLev < 4 then
			allAllowedForums = ""

			allowSql = "SELECT FORUM_ID, F_PRIVATEFORUMS, F_PASSWORD_NEW"
			allowSql = allowSql & " FROM " & strTablePrefix & "FORUM"
			allowSql = allowSql & " ORDER BY FORUM_ID"

			set rsAllowed = Server.CreateObject("ADODB.Recordset")
			rsAllowed.open allowSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rsAllowed.EOF then
				recAllowedCount = ""
			else
				allAllowedData = rsAllowed.GetRows(adGetRowsRest)
				recAllowedCount = UBound(allAllowedData,2)
			end if

			rsAllowed.close
			set rsAllowed = nothing

			if recAllowedCount <> "" then
				fFORUM_ID = 0
				fF_PRIVATEFORUMS = 1
				fF_PASSWORD_NEW = 2

				for RowCount = 0 to recAllowedCount

					Forum_ID = allAllowedData(fFORUM_ID,RowCount)
					Forum_PrivateForums = allAllowedData(fF_PRIVATEFORUMS,RowCount)
					Forum_FPasswordNew = allAllowedData(fF_PASSWORD_NEW,RowCount)

					if mLev = 4 then
						ModerateAllowed = "Y"
					elseif mLev = 3 and ModOfForums <> "" then
						if (strAuthType = "nt") then
							if (chkForumModerator(Forum_ID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
						else 
							if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") > 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
						end if
					else
						ModerateAllowed = "N"
					end if
					if ChkDisplayForum(Forum_PrivateForums,Forum_FPasswordNew,Forum_ID,MemberID) = true then
						if allAllowedForums = "" then
							allAllowedForums = Forum_ID
						else
							allAllowedForums = allAllowedForums & "," & Forum_ID
						end if
					end if
				next
			end if
			if allAllowedForums = "" then allAllowedForums = 0
		end if
        end if

	strSqlCF = "SELECT C.CAT_ID, C.CAT_NAME, F.FORUM_ID, F.F_SUBJECT, F.F_TYPE, F.F_URL"
	strSqlCF = strSqlCF & " FROM " & strTablePrefix & "FORUM F LEFT JOIN " & strTablePrefix & "CATEGORY C"
	strSqlCF = strSqlCF & " ON F.CAT_ID = C.CAT_ID"
	if strPrivateForums = "1" and allAllowedForums <> "" and mLev < 4 then
		strSqlCF = strSqlCF & " WHERE F.FORUM_ID IN (" & allAllowedForums & ")"
	end if
	strSqlCF = strSqlCF & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT DESC"

	set rsCF = Server.CreateObject("ADODB.Recordset")
	rsCF.open strSqlCF, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rsCF.EOF then
		recCatForumCount = ""
		strSelectBox = ""
	else
		allCatForumData = rsCF.GetRows(adGetRowsRest)
		recCatForumCount = UBound(allCatForumData,2)
		strSelectBox = ""
	end if

	rsCF.close
	set rsCF = nothing

	if recCatForumCount = "" then
		'Do Nothing
	else
		currForum = 0 
		cfCAT_ID = 0
		cfCAT_NAME = 1
		cfFORUM_ID = 2
		cfF_SUBJECT = 3
		cfF_TYPE = 4
		cfF_URL = 5

		for iCatForum = 0 to recCatForumCount
			CategoryID = allCatForumData(cfCAT_ID,iCatForum)
			CategoryName = allCatForumData(cfCAT_NAME,iCatForum)
			CatForumID = allCatForumData(cfFORUM_ID,iCatForum)
			CatForumSubject = allCatForumData(cfF_SUBJECT,iCatForum)
			CatForumType = allCatForumData(cfF_TYPE,iCatForum)
			CatForumURL = allCatForumData(cfF_URL,iCatForum)

			if (currCatID <> CategoryID) then
				strSelectBox = strSelectBox & "          	<option value=""default.asp?CAT_ID=" & CategoryID & """>" & CategoryName & "</option>" & vbNewline
				currCatID = CategoryID
			end if
			if CatForumType = 0 then
				strSelectBox = strSelectBox & "          	<option value=""forum.asp?FORUM_ID=" & CatForumID & """"
			elseif CatForumType = 1 then
				strSelectBox = strSelectBox & "          	<option value=""" & CatForumURL & """"
			end if
			strSelectBox = strSelectBox & ">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & CatForumSubject & "</option>" & vbNewline
		next
	end if
	Session(strCookieURL & "JumpBox") = strSelectBox
	Session(strCookieURL & "JumpBoxDate") = DateToStr(strForumTimeAdjust)
end if
Response.Write Session(strCookieURL & "JumpBox")
scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
Response.Write	"          	<option value="""">&nbsp;--------------------</option>" & vbNewline & _
		"		<option value=""" & strHomeURL & """>Home</option>" & vbNewline & _
		"		<option value=""active.asp"">Активные темы</option>" & vbNewline & _
		"		<option value=""faq.asp"">Часто задавемые вопросы</option>" & vbNewline & _
		"		<option value=""members.asp"">Информация об участниках форума</option>" & vbNewline & _
		"		<option value=""search.asp"">Поиск</option>" & vbNewline & _
		"          </select>" & vbNewline & _
		"          </form>" & vbNewline & _
		"          <!-- END JUMP TO -->" & vbNewline
%>