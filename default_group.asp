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
if strGroupCategories <> "1" then
	Response.Redirect("default.asp")
end if
if strAutoLogon = 1 then
	if (ChkAccountReg() <> "1") then
		Response.Redirect "register.asp?mode=DoIt"
	end if
end if

strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_DESCRIPTION, GROUP_ICON"
strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
strSql = strSql & " ORDER BY GROUP_NAME ASC "
set rs = my_Conn.Execute (strSql)

Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewline & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewline & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewline & _
		"              <tr>" & vbNewline & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&nbsp;</font></b></td>" & vbNewline & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Discussion&nbsp;Groups</font></b></td>" & vbNewline & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Categories</font></b></td>" & vbNewline & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Forums</font></b></td>" & vbNewline & _
		"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Topics</font></b></td>" & vbNewline & _
		"              </tr>" & vbNewline

if rs.EOF or rs.BOF then
	Response.Write	"              <tr>" & vbNewline & _
			"                <td bgcolor=""" & strCategoryCellColor & """ colspan="""
	if (strShowModerators = "1") or (mlev > 0 ) then 
		Response.Write	"6" 
	else 
		Response.Write	"5"
	end if
	Response.Write	"""><font face=""" & strDefaultFontFace & """ color=""" & strCategoryFontColor & """ size=""" & strDefaultFontSize & """ valign=""top""><b>No Categories/Forums Found</b></font></td>" & vbNewline
'	if (mlev = 4 or mlev = 3) then
        	Response.Write	"<td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ color=""" & strCategoryFontColor & """ size=""" & strDefaultFontSize & """ valign=""top"">&nbsp;</font></td>" & vbNewline
'	end if
	Response.Write	"              </tr>" & vbNewline
else
	'rs.moveFirst	           
	do until rs.EOF 
		if rs("GROUP_ID") = 1 then
			'do nothing
		else
			numCats=0
			numTopics=0
			numPosts=0
			' how many categories ?
			strSql = "SELECT GROUP_ID, GROUP_CATID "
			strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
			strSql = strSql & " WHERE GROUP_ID = " & rs("GROUP_ID")  
			strSql = strSql & " ORDER BY GROUP_ID ASC "
			set rsGroupCats = my_Conn.execute (strSql)
			if not rsGroupCats.eof then
				strSQLForum = "SELECT Count(CAT_ID) FROM " & strTablePrefix & "FORUM WHERE "
				strSQLTopic = "SELECT Count(CAT_ID) FROM " &  strTablePrefix & "TOPICS WHERE " 
				first = 0
				do until rsGroupCats.eof
					numCats = numCats + 1
					if first = 0 then 
						strSQLForum = strSQLForum & " CAT_ID =" & rsGroupCats("GROUP_CATID")
						strSQLTopic = strSQLTopic & " CAT_ID =" & rsGroupCats("GROUP_CATID")
						first = 1
					else
						strSQLForum = strSQLForum & " OR CAT_ID =" & rsGroupCats("GROUP_CATID")
						strSQLTopic = strSQLTopic & " OR CAT_ID =" & rsGroupCats("GROUP_CATID")
					end if 
					rsGroupCats.MoveNext
				loop
				rsGroupCats.close
				set rsGroupCats = nothing                  
				set rsPostCount = my_Conn.execute (strSQLTopic) 
				Select Case rsPostCount.eof
					Case False
						NumTopics = rsPostCount(0)
					Case True
						NumTopics = 0
				End Select
				set rsPostCount = nothing
				set rsGroupForums = my_Conn.execute (strSqlForum)
				Select Case rsGroupForums.eof
					Case False
						NumForums = rsGroupForums(0)
					Case True
						NumForums = 0
				End Select
				set rsGroupForums = nothing
			else 
				NumCats = 0
				NumForums = 0
				NumTopics = 0
			end if
			Response.Write	"              <tr>" & vbNewLine & _
					"                <td align=""center"" bgcolor=""" & strForumCellColor & """ nowrap valign=""top"">"
			if instr(rs("GROUP_ICON"),".") then
				Response.Write	getCurrentIcon(rs("GROUP_ICON") & "|20|20","","hspace=""0"" align=""absmiddle""") & "</td>" & vbNewLine
			else
				Response.Write	getCurrentIcon(strIconGroupCategories,"","hspace=""0"" align=""absmiddle""") & "</td>" & vbNewLine
			end if
			Response.Write	"                <td align=""left"" bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText""><a href=""default.asp?group=" & cLng(rs("GROUP_ID")) & """>" & chkString(rs("GROUP_NAME"),"display") & "</a></span>"
			if rs("GROUP_DESCRIPTION") <> "" then
				Response.Write	"<br /><font size=""" & strFooterFontSize & """>" & formatStr(rs("GROUP_DESCRIPTION")) & "</font>"
			end if
			Response.Write	"</font></td>" & vbNewLine & _
					"                <td align=""center"" bgcolor=""" & strForumCellColor & """ nowrap valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & NumCats & "</font></td>" & vbNewLine & _
					"                <td align=""center"" bgcolor=""" & strForumCellColor & """ nowrap valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & NumForums & "</font></td>" & vbNewLine & _
					"                <td align=""center"" bgcolor=""" & strForumCellColor & """ nowrap valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & NumTopics & "</font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		end if 
		rs.movenext
	loop
end if 
rs.close
set rs = nothing 
Response.Write	"            </table>" & vbNewline & _
		"          </td>" & vbNewline & _
		"        </tr>" & vbNewline & _
		"      </table><br />" & vbNewline
WriteFooter
%>