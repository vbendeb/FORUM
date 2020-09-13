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
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
intStep = Request.QueryString("Step")
if intStep = "" or IsNull(intStep) then
	intStep = 1
else
	intStep = cLng(intStep)
end if

if intStep < 5 then 
	Response.write "<meta http-equiv=""Refresh"" content=""1; URL=admin_count.asp?Step=" & intStep + 1 & """>"
else
	Response.write "<meta http-equiv=""Refresh"" content=""60; URL=admin_home.asp"">"
end if

Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Update&nbsp;Forum&nbsp;Counts<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine & _
		"      <table align=""center"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""center"" colspan=""2""><p><b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Updating Counts Step " & intStep & " of 5 </font></b><br /></td>" & vbNewLine & _
		"        </tr>" & vbNewLine
set Server2 = Server
Server2.ScriptTimeout = 6000

if intStep = 1 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top""><font face=""" &strDefaultFontFace & """>Topics:</font></td>" & vbNewline
	Response.Write "          <td valign=""top""><font face=""" &strDefaultFontFace & """>"

	'## Forum_SQL - Get contents of the Forum table related to counting
	strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recForumCount = ""
	else
		allForumData = rs.GetRows(adGetRowsRest)
		recForumCount = UBound(allForumData,2)
	end if

	rs.close
	set rs = nothing

	if recForumCount <> "" then
		fFORUM_ID = 0
		fF_TOPICS = 1
		i = 0 

		for iForum = 0 to recForumCount
			ForumID = allForumData(fFORUM_ID,iForum)
			ForumTopics = allForumData(fF_TOPICS,iForum)

			i = i + 1

			'## Forum_SQL - count total number of topics in each forum in Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & ForumID
			strSql = strSql & " AND T_STATUS <= 1 "

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_TOPICS = 0
			else
				intF_TOPICS = rs1("cnt")
			end if

			set rs1 = nothing

			'## Forum_SQL - count total number of topics in each forum in A_Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & ForumID
			strSql = strSql & " AND T_STATUS <= 1 "

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_A_TOPICS = 0
			else
				intF_A_TOPICS = rs1("cnt")
			end if

			set rs1 = nothing

			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
			strSql = strSql & " , F_A_TOPICS = " & intF_A_TOPICS
			strSql = strSql & " WHERE FORUM_ID = " & ForumID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "          <br />" & vbNewline
				i = 0
			end if
		next
	end if

	Response.Write "          </font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline

elseif intStep = 2 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top""><font face=""" &strDefaultFontFace & """>Topic Replies:</font></td>" & vbNewline
	Response.Write "          <td valign=""top""><font face=""" & strDefaultFontFace & """>"

	'## Forum_SQL
	strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS"
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_STATUS <= 1"

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recTopicCount = ""
	else
		allTopicData = rs.GetRows(adGetRowsRest)
		recTopicCount = UBound(allTopicData,2)
	end if

	rs.close
	set rs = nothing

	if recTopicCount <> "" then
		fTOPIC_ID = 0
		fT_REPLIES = 1
		i = 0 

		for iTopic = 0 to recTopicCount
			TopicID = allTopicData(fTOPIC_ID,iTopic)
			TopicReplies = allTopicData(fT_REPLIES,iTopic)

			i = i + 1

			'## Forum_SQL - count total number of replies in Topics table
			strSql = "SELECT count(REPLY_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID
			strSql = strSql & " AND R_STATUS <= 1 "

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rs.EOF then
				recReplyCntCount = ""
			else
				allReplyCntData = rs.GetRows(adGetRowsRest)
				recReplyCntCount = UBound(allReplyCntData,2)
			end if

			rs.close
			set rs = nothing

			if recReplyCntCount <> "" then
				fReplyCnt = 0

				for iCnt = 0 to recReplyCntCount
					ReplyCnt = allReplyCntData(fReplyCnt,iCnt)

					intT_REPLIES = ReplyCnt

					'## Forum_SQL - Get last_post and last_post_author for Topic
					strSql = "SELECT R_DATE, R_AUTHOR "
					strSql = strSql & " FROM " & strTablePrefix & "REPLY "
					strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
					strSql = strSql & " AND R_STATUS <= 1"
					strSql = strSql & " ORDER BY R_DATE DESC"

					set rs2 = my_Conn.Execute (strSql)

					if not(rs2.eof or rs2.bof) then
						rs2.movefirst
						strLast_Post = rs2("R_DATE")
						strLast_Post_Author = rs2("R_AUTHOR")
					else
						strLast_Post = ""
						strLast_Post_Author = ""
					end if

					set rs2 = nothing
				next
                        else
				intT_REPLIES = 0

				set rs2 = Server.CreateObject("ADODB.Recordset")

				'## Forum_SQL - Get post_date and author from Topic
				strSql = "SELECT T_AUTHOR, T_DATE "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & TopicID & " "
				strSql = strSql & " AND T_STATUS <= 1"

				set rs2 = my_Conn.Execute(strSql)

				if not(rs2.eof or rs2.bof) then
					strLast_Post = rs2("T_DATE")
					strLast_Post_Author = rs2("T_AUTHOR")
				else
					strLast_Post = ""
					strLast_Post_Author = ""
				end if

				set rs2 = nothing

			end if

			strSql = "UPDATE " & strTablePrefix & "TOPICS "
			strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
			if strLast_Post <> "" then 
				strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author 
				end if
			end if
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "          <br />" & vbNewline
				i = 0
			end if
		next
	end if

	Response.Write "          </font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline

elseif intStep = 3 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top""><font face=""" &strDefaultFontFace & """>UnModerated Topic Replies:</font></td>" & vbNewline
	Response.Write "          <td valign=""top""><font face=""" & strDefaultFontFace & """>"

	'## Forum_SQL
	strSql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS"
	strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.T_STATUS <= 1"

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recTopicCount = ""
	else
		allTopicData = rs.GetRows(adGetRowsRest)
		recTopicCount = UBound(allTopicData,2)
	end if

	rs.close
	set rs = nothing

	if recTopicCount <> "" then
		fTOPIC_ID = 0
		i = 0 

		for iTopic = 0 to recTopicCount
			TopicID = allTopicData(fTOPIC_ID,iTopic)

			i = i + 1

			'## Forum_SQL - count total number of unmoderated replies in Topics table
			strSql = "SELECT count(REPLY_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID
			strSql = strSql & " AND R_STATUS = 2 OR R_STATUS = 3 "

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rs.EOF then
				recReplyCntCount = ""
			else
				allReplyCntData = rs.GetRows(adGetRowsRest)
				recReplyCntCount = UBound(allReplyCntData,2)
			end if

			rs.close
			set rs = nothing

			if recReplyCntCount <> "" then
				fReplyCnt = 0
				for iCnt = 0 to recReplyCntCount
					intT_UREPLIES = allReplyCntData(fReplyCnt,iCnt)
				next
                        else
				intT_UREPLIES = 0
			end if

			strSql = "UPDATE " & strTablePrefix & "TOPICS "
			strSql = strSql & " SET T_UREPLIES = " & intT_UREPLIES
			strSql = strSql & " WHERE TOPIC_ID = " & TopicID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "          <br />" & vbNewline
				i = 0
			end if
		next
	end if

	Response.Write "          </font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline

elseif intStep = 4 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top""><font face=""" & strDefaultFontFace & """>Forum Replies:</font></td>" & vbNewline
	Response.Write "          <td valign=top><font face=""" &strDefaultFontFace & """>"

	'## Forum_SQL - Get values from Forum table needed to count replies
	strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rs.EOF then
		recForumCount = ""
	else
		allForumData = rs.GetRows(adGetRowsRest)
		recForumCount = UBound(allForumData,2)
	end if

	rs.close
	set rs = nothing

	if recForumCount <> "" then
		fFORUM_ID = 0
		fF_COUNT = 1
		i = 0

		for iForum = 0 to recForumCount
			ForumID = allForumData(fFORUM_ID,iForum)
			ForumCount = allForumData(fF_COUNT,iForum)

			i = i + 1

			'## Forum_SQL - Count total number of Replies
			strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & ForumID
			strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_COUNT = 0
				intF_TOPICS = 0
			else
				intF_COUNT = rs1("cnt") + rs1("SumOfT_REPLIES")
				intF_TOPICS = rs1("cnt") 
			end if
			if IsNull(intF_COUNT) then intF_COUNT = 0 
			if IsNull(intF_TOPICS) then intF_TOPICS = 0 

			set rs1 = nothing

			'## Forum_SQL - Count total number of Archived Replies
			strSql = "SELECT Sum(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
			strSql = strSql & " WHERE " & strTablePrefix & "A_TOPICS.FORUM_ID = " & ForumID
			strSql = strSql & " AND " & strTablePrefix & "A_TOPICS.T_STATUS <= 1"

			set rs1 = my_Conn.Execute(strSql)

			if rs1.EOF or rs1.BOF then
				intF_A_COUNT = 0
				intF_A_TOPICS = 0
			else
				intF_A_COUNT = rs1("cnt") + rs1("SumOfT_REPLIES")
				intF_A_TOPICS = rs1("cnt") 
			end if
			if IsNull(intF_A_COUNT) then intF_A_COUNT = 0 
			if IsNull(intF_A_TOPICS) then intF_A_TOPICS = 0 

			set rs1 = nothing

			'## Forum_SQL - Get last_post and last_post_author for Forum
			strSql = "SELECT T_LAST_POST, T_LAST_POST_AUTHOR "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & ForumID & " "
			strSql = strSql & " AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"
			strSql = strSql & " ORDER BY T_LAST_POST DESC"

			set rs2 = my_Conn.Execute (strSql)

			if not (rs2.eof or rs2.bof) then
				strLast_Post = rs2("T_LAST_POST")
				strLast_Post_Author = rs2("T_LAST_POST_AUTHOR")
			else
				strLast_Post = ""
				strLast_Post_Author = ""
			end if

			set rs2 = nothing

			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_COUNT = " & intF_COUNT
			strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
			strSql = strSql & ",  F_A_COUNT = " & intF_A_COUNT
			strSql = strSql & ",  F_A_TOPICS = " & intF_A_TOPICS
			if strLast_Post <> "" then 
				strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "' "
				if strLast_Post_Author <> "" then 
					strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
				end if
			end if
			strSql = strSql & " WHERE FORUM_ID = " & ForumID

			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords

			Response.Write "."
			if i = 80 then 
				Response.Write "          <br />" & vbNewline
				i = 0
			end if	
		next
	end if
	Response.Write "          </font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline

elseif intStep = 5 then 

	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""right"" valign=""top""><font face=""" &strDefaultFontFace & """>Totals:</font></td>" & vbNewline
	Response.Write "          <td valign=""top""><font face=""" &strDefaultFontFace & """>"

	'## Forum_SQL - Total of Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
	strSql = strSql & " AS SumOfF_TOPICS "
	strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_TOPICS) "
	strSql = strSql & " AS SumOfF_A_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = my_Conn.Execute(strSql)

	if rs("SumOfF_TOPICS") <> "" then
		Response.Write "Total Topics: " & rs("SumOfF_TOPICS") & "<br />" & vbNewline
		intSumOfF_TOPICS = rs("SumOfF_TOPICS")
	else
		Response.Write "Total Topics: 0<br />" & vbNewLine
		intSumOfF_TOPICS = 0
	end if
	if rs("SumOfF_A_TOPICS") <> "" then
		Response.Write "Archived Topics: " & rs("SumOfF_A_TOPICS") & "<br />" & vbNewline
		intSumOfF_A_TOPICS = rs("SumOfF_A_TOPICS")
	else
		Response.Write "Archived Topics: 0<br />" & vbNewLine
		intSumOfF_A_TOPICS = 0
	end if

	'## Forum_SQL - Write total Topics to Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET T_COUNT = " & intSumOfF_TOPICS
	strSql = strSql & " , T_A_COUNT = " & intSumOfF_A_TOPICS

	set rs = nothing

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	'## Forum_SQL - Total all the replies for each topic
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
	strSql = strSql & " AS SumOfF_COUNT "
	strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_COUNT) "
	strSql = strSql & " AS SumOfF_A_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "

	set rs = my_Conn.Execute (strSql)

	if rs("SumOfF_COUNT") <> "" then
		Response.Write "          Total Posts: " & RS("SumOfF_COUNT") & "<br />" & vbNewline
		intSumOfF_COUNT = rs("SumOfF_COUNT")
	else
		Response.Write "          Total Posts: 0<br />" & vbNewline
		intSumOfF_COUNT = 0
	end if
	if rs("SumOfF_A_COUNT") <> "" then
		Response.Write "          Total Archived Posts: " & RS("SumOfF_A_COUNT") & "<br />" & vbNewline
		intSumOfF_A_COUNT = rs("SumOfF_A_COUNT")
	else
		Response.Write "          Total Posts: 0<br />" & vbNewline
		intSumOfF_A_COUNT = 0
	end if

	'## Forum_SQL - Write total replies to the Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET P_COUNT = " & intSumOfF_COUNT
	strSql = strSql & " , P_A_COUNT = " & intSumOfF_A_COUNT

	set rs = nothing

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	'## Forum_SQL - Total number of users
	strSql = "SELECT Count(MEMBER_ID) "
	strSql = strSql & " AS CountOf "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"

	set rs = my_Conn.Execute(strSql)

	Response.Write "          Registered Users: " & rs("Countof") & "<br />" & vbNewline

	'## Forum_SQL - Write total number of users to Totals table
	strSql = " UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET U_COUNT = " & cLng(RS("Countof"))

	set rs = nothing

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	Response.Write "          </font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline
	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""center"" colspan=""2"">&nbsp;<br />" & vbNewline
	Response.Write "          <b><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Count Update Complete</font></b></font></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline
	Response.Write "        <tr>" & vbNewline
	Response.Write "          <td align=""center"" colspan=""2"">&nbsp;<br />" & vbNewline
	Response.Write "          <a href=""admin_home.asp""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strDefaultFontColor & """>Back to Admin Home</font></a></td>" & vbNewline
	Response.Write "        </tr>" & vbNewline
end if


response.write "      </table>"

WriteFooter
Response.End
%>
