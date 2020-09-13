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
Forum_ID = Request.QueryString("Forum")
User_ID = Request.QueryString("userid")
Action_ID = Request.QueryString("action")
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp"-->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Moderator Configuration<br /><br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if Forum_ID = "" then
	txtMessage = "Select a forum to edit moderators for that forum"
else
	if User_ID = "" then
		txtMessage = "Select a user to grant/revoke moderator powers for that user.	Users in bold are currently moderators of this forum."
	else
		if Action_ID = "" then
			txtMessage = "Select an action for this user"
		else
			txtMessage = "Action Successful"
		end if
	end if
end if

Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor  & """><b>Moderator Configuration</b>"
if txtMessage <> "" Then Response.Write("<br />" & txtMessage)
Response.Write	"</font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine
if Forum_ID = "" then
	Response.Write	"                  <table>" & vbNewLine

	'## Forum_SQL
	strSql = "SELECT C.CAT_ORDER, C.CAT_NAME, F.CAT_ID, F.FORUM_ID, F.F_ORDER, F.F_SUBJECT " &_
	" FROM " & strTablePrefix & "CATEGORY C, " & strTablePrefix & "FORUM F" &_
	" WHERE C.CAT_ID = F.CAT_ID "
	strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT ASC;"

	set rs = my_Conn.Execute(strSql)

	if rs.eof or rs.bof then
		'nothing
	else
		iOldCat = 0
		do until rs.EOF
			iNewCat = rs("CAT_ID")
			if iNewCat <> iOldCat Then
				Response.Write	"                    <tr><td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><b>" & rs("CAT_NAME") & "</b></font></td></tr>" & vbNewLine
				iOldCat = iNewCat
			end if
			Response.Write	"                    <tr><td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>&nbsp;&nbsp;&nbsp;&nbsp;<span class=""spnMessageText""><a href=""admin_moderators.asp?forum=" & rs("FORUM_ID") & """>" & rs("F_SUBJECT") & "</a></span></font></td></tr>" & vbNewLine
			rs.MoveNext
		loop
	end if
	Response.Write	"                  </table>" & vbNewLine
else
	if Action_ID = "" then
		if User_ID = "" then

			'## Forum_SQL
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
			strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
			strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"

			set rs = my_Conn.Execute(strSql)

			Response.Write	"                <br />" & vbNewLine & _
					"                <ul>" & vbNewLine
			do until rs.EOF
				Response.Write	"                <li>"
				if chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("<b>")
				Response.Write	"<span class=""spnMessageText""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & rs("MEMBER_ID")& """>" & rs("M_NAME") & "</a></span>"
				If chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("</b>")
				Response.Write	"</li>" & vbNewLine
				rs.MoveNext
			loop
			Response.Write	"                </ul>" & vbNewLine
		else

			'## Forum_SQL
			strSql = "SELECT " & strTablePrefix & "MODERATOR.FORUM_ID, " & strTablePrefix & "MODERATOR.MEMBER_ID, " & strTablePrefix & "MODERATOR.MOD_TYPE "
			strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
			strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID & " "
			strSql = strSql & " AND " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "

			set rs = my_Conn.Execute(strSql)

			if rs.EOF then
				Response.Write	"                <center>" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                The selected user is not a moderator of the selected forum<br />" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                If you would like to make this user the moderator of this forum, <span class=""spnMessageText""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=1"">click here</a></span>." & vbNewLine & _
						"                </center>" & vbNewLine & _
						"                <br />" & vbNewLine
			else
				Response.Write	"                <center>" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                The selected user is currently a moderator of the selected forum<br />" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                If you would like to remove this user's moderator status in this forum, <span class=""spnMessageText""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=2"">click here</a></span>." & vbNewLine & _
						"                </center>" & vbNewLine & _
						"                <br />" & vbNewLine
			end if
		end if
	else
		select case Action_ID
			case 1
				'## Forum_SQL
				strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
				strSql = strSql & "(FORUM_ID"
				strSql = strSql & ", MEMBER_ID"
				strSql = strSql & ") VALUES (" 
				strSql = strSql & Forum_ID
				strSql = strSql & ", " & User_ID
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"                <br />" & vbNewLine & _
						"                <center>" & vbNewLine & _
						"                The selected user is now a moderator of the selected forum<br />" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                <span class=""spnMessageText""><a href=""admin_moderators.asp"">Back to Moderator Options</a></span>" & vbNewLine & _
						"                </center><br />" & vbNewLine
			case 2

				'## Forum_SQL
				strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
				strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "
				strSql = strSql & " AND   " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write	"                <br />" & vbNewLine & _
						"                <center>" & vbNewLine & _
						"                The selected user's moderator status in the selected forum has been removed<br />" & vbNewLine & _
						"                <br />" & vbNewLine & _
						"                <span class=""spnMessageText""><a href=""admin_moderators.asp"">Back to Moderator Options</a></span>" & vbNewLine & _
						"                </center><br />" & vbNewLine
		end select
	end if
end if
Response.Write	"                </font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table><br />" & vbNewLine
WriteFooter
Response.End
%>
