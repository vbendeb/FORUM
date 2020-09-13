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
<%
dim Subscribe, sublevel, CatID, ForumID, TopicID, RecordCount, ThisMemberID

ThisMemberID = cLng(Request("MEMBER_ID"))

if MemberID < 0 then
	Response.Write("<script language=""JavaScript"" type=""text/javascript"">this.window.close();</script>") & vbNewLine
	Response.End
end if

if (ThisMemberID <> MemberID and mlev = 4) or (MemberID = ThisMemberID) then
	Subscribe = Request.QueryString("SUBSCRIBE")
	SubLevel  = Request.QueryString("LEVEL")
	CatID     = Request.QueryString("CAT_ID")
	if CatID = "" then CatId = 0 else CatID = cLng(Request.QueryString("CAT_ID"))
	ForumID   = Request.QueryString("FORUM_ID")
	if ForumID = "" then ForumId = 0 else ForumID = cLng(Request.QueryString("FORUM_ID"))
	TopicID   = Request.QueryString("TOPIC_ID")
	if TopicID = "" then TopicId = 0 else TopicID = cLng(Request.QueryString("TOPIC_ID"))
	Member_ID  = cLng(Request.QueryString("MEMBER_ID"))

	Response.Write	"      <p align=center>" & vbNewline
	Response.Write	"      <font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>" & vbNewline

	' --- Is the member trying to subscribe or unsubscribe??
	Select case Subscribe
			' --- Unsubscribe
	         case "U"
		        DeleteSubscription sublevel, Member_ID, CatID, ForumID, TopicID
			' --- Return the appropriate message to the user....
			if CheckSubscriptionCount(SubLevel, Member_ID, CatID, ForumID, TopicID) > 0 then
				Response.Write	"      <b>Subscriptions"
	                 else
		         	Response.Write	"      <b>Subscription"
			 end if
				Response.Write	" Cancelled!</b></font></p>" & vbNewline & _
						"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
	         ' --- Subscribe
		     case "S"
			' --- Check for overriding subscriptions to prevent duplicate emails
			if (sublevel = "TOPIC" or sublevel = "FORUM" or sublevel = "CAT") and (CheckSubscriptionCount("BOARD", Member_ID, 0, 0, 0) > 0) then
				SendHigherLevelMsg "BOARD", 0
			elseif (sublevel = "TOPIC" or sublevel = "FORUM") and (CheckSubscriptionCount("CAT", Member_ID, CatId, 0, 0) > 0) then
				SendHigherLevelMsg "CAT", CatId
			elseif sublevel = "TOPIC"  and (CheckSubscriptionCount("FORUM", Member_ID, CatID, ForumID, 0) > 0) then
				SendHigherLevelMsg "FORUM", ForumId
			else
				' Delete any lower subscriptions to prevent duplicates emails.....
				if SubLevel = "FORUM" or SubLevel = "CAT" or SubLevel = "BOARD" then
					DeleteSubscription sublevel, Member_ID, CatID, ForumID, TopicID
				end if
				AddSubscription SubLevel, Member_ID, CatID, ForumID, TopicID
	                 end if
	End Select
else
	Response.Write	"      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>You do not have permission to change another users subscription.  Only Administrators may change another users subscriptions.</font></p>" & vbNewline

	' ## This is just the form which is used to login if the person is
	' ## not logged in or does not have access to do the moderation.
	Response.Write	"      <form action=""pop_subscription.asp?UserCheck=Y"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewline & _
			"      <input type=""hidden"" name=""REPLY_ID"" value=""" & ReplyID & """>" & vbNewline & _
			"      <input type=""hidden"" name=""TOPIC_ID"" value=""" & TopicID & """>" & vbNewline & _
			"      <input type=""hidden"" name=""FORUM_ID"" value=""" & ForumID & """>" & vbNewline & _
			"      <input type=""hidden"" name=""CAT_ID""   value=""" & CatID & """>"   & vbNewline & _
			"      <table border=""0"" width=""75%"" cellspacing=""0"" cellpadding=""0"">" & vbNewline & _
			"        <tr>" & vbNewline & _
			"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewline & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewline
	if strAuthType = "db" then
		Response.Write	"              <tr>" & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>User Name:</font></b></td>" & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""name"" value=""" & strDBNTUserName & """ size=""20""></td>" & vbNewline & _
				"              </tr>" & vbNewline & _
				"              <tr>"  & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Password:</font></b></td>" & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """><input type=""password"" name=""password"" value="""" size=""20""></td>" & vbNewline & _
				"              </tr>" & vbNewline
	else
		Response.Write	"              <tr>"  & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>NT Account:</font></b></td>" & vbNewline & _
				"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""DBNTUserName"" value=""" & strDBNTUserName & """ size=""20""></td>" & vbNewline & _
				"              </tr>" & vbNewline
	end if
	Response.Write	"              <tr>" & vbNewline & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><Input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & vbNewline & _
			"              </tr>" & vbNewline & _
			"            </table>"  & vbNewline & _
			"          </td>" & vbNewline & _
			"        </tr>" & vbNewline & _
			"      </table>" & vbNewline & _
			"      </form>" & vbNewline
end if

WriteFooterShort
Response.End

sub DeleteSubscription(Level, Member_ID, CatID, ForumID, TopicID)
	' --- Delete the appropriate sublevel of subscriptions
 	StrSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS"
  	StrSql = StrSql & " WHERE " & strTablePrefix & "SUBSCRIPTIONS.MEMBER_ID = " & Member_ID
   	if sublevel = "CAT" then
		strSql = strSql & " AND " & strTablePrefix & "SUBSCRIPTIONS.CAT_ID = " & CatID
	elseif sublevel = "FORUM" then
		strSql = strSql & " AND " & strTablePrefix & "SUBSCRIPTIONS.FORUM_ID = " & ForumID
	elseif sublevel = "TOPIC" then
		strSql = strSql & " AND " & strTablePrefix & "SUBSCRIPTIONS.TOPIC_ID = " & TopicID
	end if
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

sub AddSubscription(SubLevel, Member_ID, CatID, ForumID, TopicID)
	' --- Insert the appropriate sublevel subscription
        strSql = "INSERT INTO " & strTablePrefix & "SUBSCRIPTIONS"
        strSql = strSql & "(MEMBER_ID, CAT_ID, FORUM_ID, TOPIC_ID) VALUES (" & Member_ID & ", "
        if sublevel = "BOARD" then
		strSql = strSql & "0, 0, 0)"
	elseif sublevel = "CAT" then
		strSql = strSql & CatID & ", 0, 0)"
	elseif sublevel = "FORUM" then
		strSql = strSql & CatID & ", " & ForumID & ", 0)"
	else
		strSql = strSql & CatID & ", " & ForumID & ", " & TopicID & ")"
	end if
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	Response.Write	"You are subscribed to "
	if sublevel = "BOARD" then
		Response.Write	"<br />all posts in the " 
		Response.Write	"<br />" & strForumTitle & " forums "
	elseif sublevel = "CAT" then
		strSql = "SELECT " & strTablePrefix & "CATEGORY.CAT_NAME "
		strSql = strSql & "FROM " & strTablePrefix & "CATEGORY "
		strSql = strSql & "WHERE " & strTablePrefix & "CATEGORY.CAT_ID = " & CatID

		set rs = my_Conn.Execute (strSql)
		strCategory = rs("CAT_NAME")

		rs.close
		set rs = nothing

		Response.Write	"<br />all posts in "
		Response.Write	"<br />" & strCategory

	elseif sublevel = "FORUM" then
		strSql = "SELECT " & strTablePrefix & "FORUM.F_SUBJECT "
		strSql = strSql & "FROM " & strTablePrefix & "FORUM "
		strSql = strSql & "WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & ForumId

		set rs = my_Conn.Execute (strSql)
		strForum = rs("F_SUBJECT")

		rs.close
		set rs = nothing

		Response.Write	"<br />all posts in "
		Response.Write	"<br />" & strForum
	else
		Response.Write	"<br />all replies made to this Topic"
	end if
	Response.Write	"</font></p>" & vbNewline & _
			"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
end sub

sub SendHigherLevelMsg(SubLevel, Id)
	' -- If an overriding subscription is found, return the appropriate error message.
	dim rs
	Response.Write	"You currently are subscribed to "
	if sublevel = "BOARD" then
		Response.Write "<br />all posts in " & strForumTitle & ""
	elseif sublevel = "CAT" then
		strSql = "SELECT CAT_NAME "
		strSql = strSql & "FROM " & strTablePrefix & "CATEGORY "
		strSql = strSql & "WHERE CAT_ID = " & Id

	        set rs = my_Conn.Execute (strSql)
		strCategory = rs("CAT_NAME")

		rs.close
		set rs = nothing

		Response.Write	"<br />all posts in "
		Response.Write	"<br />" & strCategory

	elseif sublevel = "FORUM" then
		strSql = "SELECT F_SUBJECT "
		strSql = strSql & "FROM " & strTablePrefix & "FORUM "
		strSql = strSql & "WHERE FORUM_ID = " & Id

	        set rs = my_Conn.Execute (strSql)
		strForum = rs("F_SUBJECT")

		rs.close
		set rs = nothing

		Response.Write	"<br />all posts in "
		Response.Write	"<br />" & strForum

	end if
	Response.Write	"<br />This will also mail you notification at the level you requested.</font></p>" & vbNewline & _
			"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
end sub

function CheckSubscriptionCount(Level, Member_ID, CatID, ForumID, TopicID)
	' --- Count the number of subscriptions at the appropriate sublevel.
	dim SubCount
	StrSql = "SELECT Count(*) as RecordCount from " & strTablePrefix & "SUBSCRIPTIONS S"
	StrSql = StrSql & " WHERE S.MEMBER_ID = " & Member_ID
	if Level = "CAT" then
		StrSql = StrSQL & " AND S.CAT_ID = " & CatID
	elseif Level = "FORUM" then
		StrSql = StrSQL & " AND S.FORUM_ID = " & ForumID
	elseif Level = "TOPIC" then
		StrSql = StrSQL & " AND S.TOPIC_ID = " & TopicID
	else ' BOARD-level
		StrSql = StrSQL & " AND S.CAT_ID = 0 "
		StrSql = StrSQL & " AND S.FORUM_ID = 0 "
		StrSql = StrSQL & " AND S.TOPIC_ID = 0 "
	end if
	set rs1 = my_Conn.Execute (strSql)
	if rs1.EOF or rs1.BOF then
		SubCount = 0
	else
		SubCount = rs1("RecordCount")
	end if
	rs1.Close
	set rs1 = nothing
	CheckSubscription = SubCount
end function
%>