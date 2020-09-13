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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<% 
Topic_ID = cLng(Request.QueryString("TOPIC_ID"))
if Topic_ID = 0 then
	Go_Result "Topic not found."
	Response.End
end if	

if Request.QueryString("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
else
	strActivePrefix = strTablePrefix
end if

'## Forum_SQL - Get Origional Posting
strSql = "SELECT M.M_NAME, M.MEMBER_ID, T.T_DATE, T.T_SUBJECT, T.T_AUTHOR, T.FORUM_ID, T.TOPIC_ID, T.T_MESSAGE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "TOPICS T"
strSql = strSql & " WHERE M.MEMBER_ID = T.T_AUTHOR "
strSql = strSql & " AND T.T_STATUS < " & 2
strSql = strSql & " AND T.TOPIC_ID = " &  Topic_ID 

set rs4 = my_Conn.Execute (strSql)
if rs4.EOF then
	rs4.close
	set rs4 = nothing
	Go_Result "Either the Topic was not found<br />or you are not authorized to view it."
	Response.End
end if

Forum_ID = rs4("FORUM_ID")
if strPrivateForums = "1" then
	result = chkForumAccess(Forum_ID,MemberID,false)
	if result = "False" or result = "FALSE" then
		Go_Result "You do not have access to<br />the forum where this Topic resides."
		Response.End
	end if
end if

'## Forum_SQL - Get all replies to this topic from DB
strSql = "SELECT M.M_NAME, R.REPLY_ID, R.R_AUTHOR, R.TOPIC_ID, R.R_DATE, R.R_MESSAGE "
strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "REPLY R "
strSql = strSql & " WHERE M.MEMBER_ID = R.R_AUTHOR "
strSql = strSql & " AND R_STATUS < " & 2
strSql = strSql & " AND R.TOPIC_ID = " & Topic_ID
strSql = strSql & " ORDER BY R.R_DATE"

set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.open  strSql, my_Conn		

Response.Write	"    <a href=""javascript:onClick=window.print()"">Print Page</a> | <a href=""JavaScript:onClick=window.close()"">Close Window</a></font><br />" & vbNewline & _
		"    </div></center>" & vbNewline & _
		"    </td>" & vbNewline & _
		"  </tr>" & vbNewline & _
		"  <tr>" & vbNewline & _
		"    <td align=""left"" valign=""top"">" & vbNewline & _
		"    <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewline &_
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>" & chkString(rs4("T_SUBJECT"),"display") & "</b></p>" & vbNewline & _
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Printed from:</b> " & strForumTitle & "<br />" & vbNewline & _
		"    <b>Topic URL:</b> <a href=""" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & """ target=""_blank"">" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "</a><br />" & vbNewline & _
		"    <b>Printed on:</b> " & ChkDate(DateToStr(Now()),"",false) & "</p>" & vbNewline & _
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Topic: </b><br />" & vbNewline & _
		"    <hr></p></div align=""center""></center>" & vbNewline & _
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Topic author:</b> " & chkString(rs4("M_NAME"),"display") & "</br>" & vbNewline & _
		"    <b>Subject:</b> " & chkString(rs4("T_SUBJECT"),"display") & "<br />" & vbNewline & _
		"    <b>Posted on:</b> " & ChkDate(rs4("T_DATE"), " " ,true) & "<br />" & vbNewline & _
		"    <b>Message:</b><br /><p align=""left"">" & formatStr(rs4("T_MESSAGE")) & "</p>" & vbNewline

if rs3.EOF or rs3.BOF then  
	'## Do Nothing
else
	rs3.movefirst
	Response.Write	"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>" & vbNewline & _
			"    <b>Replies: </b></font><br />" & vbNewline
	do until rs3.EOF
		Response.Write	"    <hr></p>" & vbNewline & _
				"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewline & _
				"    <b>Reply author:</b> " & chkString(rs3("M_NAME"),"display") & "<br />" & vbNewline & _
				"    <b>Replied on:</b> " & ChkDate(rs3("R_DATE"), " " ,true) & "<br />" & vbNewline & _
				"    <b>Message:</b></p><p align=""left"">" & formatStr(rs3("R_MESSAGE")) & "</p>" & vbNewline
		rs3.MoveNext
	loop
end if

rs3.close
set rs3 = Nothing
rs4.close
set rs4 = Nothing

Response.Write	"    <p align=""left""><hr></p>" & vbNewline & _
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>" & strForumTitle & " </b>: <a href=""" & strForumURL & """>" & strForumURL & "</a></p>" & vbNewline & _
		"    <p align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>&copy; " & strCopyright & "</b></font></p>" & vbNewline &_
		"    </td>" & vbNewline & _
		"  </tr>" & vbNewline & _
		"  <tr>" & vbNewline & _
		"    <td align=""center"" valign=""middle"">" & vbNewline & _
		"    <div align=""center""><center>" & vbNewline & _
		"    <font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewline
WriteFooterShort
Response.End

function Go_Result(message)
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """><b>There has been a problem!</b></font></p>" & vbNewLine & _
			"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """>" & message & "</font></p>" & vbNewLine
	WriteFooterShort
	Response.End
end function
%>