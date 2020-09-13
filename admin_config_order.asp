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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	Response.Redirect "admin_login_short.asp?target=admin_config_order.asp"
end if
if Request.Form("Method_Type") = "Write_Configuration" then 
	if Request.Form("NumberCategories") <> "" then
		i = 1
		do until i > cLng(Request.Form("NumberCategories"))
			SelectName = Request.Form("SortCategory" & i)
			if isNull(SelectName) then SelectName = cLng(Request.Form("NumberCategories"))
			SelectID = Request.Form("SortCatID" & i)
			NumberForums = Request.Form("NumberForums" & SelectID)
			
			'## Forum_SQL - Do DB Update
			strSql = "UPDATE " & strTablePrefix & "CATEGORY "
			strSql = strSql & " SET CAT_ORDER = " & SelectName & " "
			strSql = strSql & " WHERE CAT_ID = " & SelectID & " "
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			if NumberForums <> "" then
				j = 1
				do until j > cLng(Request.Form("NumberForums" & SelectID))
					SelectNamec = Request.Form("SortCat" & i & "SortForum" & j)
					if isNull(SelectNamec) then SelectNamec = cLng(Request.Form("NumberForums" & SelectID))
					SelectIDc = Request.Form("SortCatID" & i & "SortForumID" & j)
			
					'## Forum_SQL - Do DB Update
					strSql = "UPDATE " & strTablePrefix & "FORUM "
					strSql = strSql & " SET F_ORDER = " & SelectNamec & " "
					strSql = strSql & " WHERE FORUM_ID = " & SelectIDc & " "
					strSql = strSql & " AND CAT_ID = " & SelectID & " "

					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

					j = j + 1
				loop
			end if
			i = i + 1
		loop
	end if
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Order Submitted!</font></p>" & vbNewline & _
			"      <script language=""javascript1.2"">self.opener.location.reload();</script>" & vbNewLine
	'<meta http-equiv="Refresh" content="2; URL=admin_home.asp">
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewline
else
	Response.Write	"      " & strParagraphFormat1 & "<b>Category/Forum Order Configuration</b></font></p>" & vbNewLine
	Response.Write	"      <form action=""admin_config_order.asp"" method=""post"" id=""Form1"" name=""Form1"">" & vbNewline & _
			"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine
	'## Forum_SQL - Get all Forums From DB
	strSql = "SELECT CAT_ID, CAT_STATUS, CAT_NAME, CAT_ORDER "
	strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
	strSql = strSql & " ORDER BY CAT_ORDER, CAT_NAME "

	set rs = Server.CreateObject("ADODB.Recordset")

	if strDBType = "mysql" then
		'## Forum_SQL
		strSql2 = "SELECT COUNT(CAT_ID) AS PAGECOUNT "
		strSql2 = strSql2 & " FROM " & strTablePrefix & "CATEGORY" 
		set rsCount = my_Conn.Execute(strSql2)
		categorycount = rsCount("PAGECOUNT")
		rsCount.close

		rs.open strSql, my_Conn, adOpenStatic
	else
		rs.cachesize = 20
		rs.open strSql, my_Conn, adOpenStatic

		if not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize = 1
			categorycount = cLng(rs.pagecount)
		end if
	end if
	Response.Write	"      <input name=""NumberCategories"" type=""hidden"" value=""" & categorycount & """>"  & vbNewline
	
	Response.Write	"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewline & _
			"        <tr>" & vbNewline & _ 
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewline & _
			"            <table border=""0"" cellspacing=""1"" cellpadding=""4"">" & vbNewline & _
			"              <tr>" & vbNewline & _
			"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Category/Forum</font></b></td>" & vbNewline & _
			"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ nowrap valign=""top""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Order</font></b></td>" & vbNewline & _
			"              </tr>" & vbNewline
	if rs.EOF or rs.BOF then
		Response.Write	"              <tr>" & vbNewline & _
				"                <td bgcolor=""" & strCategoryCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ color=""" & strCategoryFontColor & """ size=""" & strDefaultFontSize & """ valign=""top""><b>No Categories/Forums Found</b></font></td>" & vbNewline & _
				"              </tr>" & vbNewline
	else
		catordercount = 1
		do until rs.EOF 
			'## Forum_SQL - Build SQL to get forums via category
			strSql = "SELECT FORUM_ID, F_SUBJECT, CAT_ID, F_TYPE, F_ORDER "
			strSql = strSql & "FROM " & strTablePrefix & "FORUM"
			strSql = strSql & " WHERE CAT_ID = " & rs("CAT_ID")
			strSql = strSql & " ORDER BY F_ORDER ASC, F_SUBJECT ASC;"
			set rsForum = Server.CreateObject("ADODB.Recordset")
			rsForum.open strSql, my_Conn, adOpenStatic

			if NOT (rsForum.EOF or rsForum.BOF) then
				rsForum.movefirst
				rsForum.pagesize = 1
			end if
			if strDBType = "mysql" then
				'## Forum_SQL
				strSql2 = "SELECT COUNT(F.FORUM_ID) AS PAGECOUNT "
				strSql2 = strSql2 & " FROM " & strTablePrefix & "FORUM F" 
				strSql2 = strSql2 & " WHERE F.CAT_ID = " & rs("CAT_ID")
				set rsCount = my_Conn.Execute(strSql2)
				forumcount = rsCount("PAGECOUNT")
				rsCount.close
				set rsCount = nothing
			else
				forumcount = cLng(rsForum.pagecount)
			end if

			Response.Write"<input name=""NumberForums" & rs("CAT_ID") & """ type=""hidden"" value=""" & forumcount & """> " & vbNewline
			chkDisplayHeader = true
			if rsForum.eof or rsForum.bof then
				Response.Write	"              <tr>" & vbNewLine & _
						"                <td bgcolor=""" & strCategoryCellColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>" & ChkString(rs("CAT_NAME"),"display") & "</b></font></td>" & vbNewline & _
						"                <td bgcolor=""" & strCategoryCellColor & """ align=""center"">" & vbNewLine
				SelectName = "SortCategory" & catordercount
				SelectID   = "SortCatID" & catordercount
				Response.Write	"                <input name=""" & SelectID & """ type=""hidden"" value=""" & rs("CAT_ID") & """>" & vbNewline & _
						"                <select name=""" & SelectName & """>" & vbNewline
				i = 1
				do while i <= categorycount
					Response.Write "                 	<option value=""" & i & """" & chkSelect(i,rs("CAT_ORDER")) & ">" & i & "</option>" & vbNewline
					i = i + 1
				loop 
				Response.Write	"                </select></td>" & vbNewline & _
						"              </tr>" & vbNewline & _
						"              <tr>" & vbNewline & _
						"                <td colspan=""2"" bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """><b>No Forums Found</b></font></td>" & vbNewline & _
						"      	       </tr>" & vbNewline
			else
				forumordercount = 1
				do until rsForum.Eof
					if rsForum("F_TYPE") <> "1" then 
						intForumCount = intForumCount + 1
					end if
					if chkDisplayHeader then
					Response.Write	"              <tr>" & vbNewline & _
								"                <td bgcolor=""" & strCategoryCellColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>" & ChkString(rs("CAT_NAME"),"display") & "</b></font></td>" & vbNewline & _
								"                <td bgcolor=""" & strCategoryCellColor & """ align=""center"">" & vbNewLine
						SelectName = "SortCategory" & catordercount
						SelectID = "SortCatID" & catordercount
						Response.Write	"                <input name=""" & SelectID & """ type=""hidden"" value=""" & rs("CAT_ID") & """>" & vbNewline & _
								"                <select name=""" & SelectName & """>" & vbNewline
						i = 1
						do while i <= categorycount
							Response.Write	"                	<option value=""" & i & """" & chkSelect(i,rs("CAT_ORDER")) & ">" & i & "</option>" & vbNewline
							i = i + 1
						loop 
						Response.Write	"                </select></td>" & vbNewline & _
								"              </tr>" & vbNewline
						chkDisplayHeader = false
					end if
					if rsForum("F_TYPE") = "1" then strType = getCurrentIcon(strIconUrl,"Web Link","hspace=""0"" align=""absmiddle""") else strType = getCurrentIcon(strIconBlank,"","hspace=""0"" align=""absmiddle""")
					Response.Write	"              <tr>" & vbNewline & _
							"                <td bgcolor=""" & strForumCellColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><b>" & strType & "&nbsp;" & ChkString(rsForum("F_SUBJECT"),"display") & "</b></font></td>" & vbNewline & _
							"                <td bgcolor=""" & strForumCellColor & """ align=""center"">" & vbNewline
					SelectName = "SortCat" & catordercount & "SortForum" & forumordercount
				        SelectID   = "SortCatID" & catordercount & "SortForumID" & forumordercount
					Response.Write	"                <input name=""" & SelectID & """ type=""hidden"" value=""" & rsForum("FORUM_ID") & """>" & vbNewline & _
							"                <select name=""" & SelectName & """>" & vbNewline
				        i = 1
		        	 	do while i <= forumcount
						Response.Write	"                	<option value=""" & i & """" & chkSelect(i,rsForum("F_ORDER")) & ">" & i & "</option>" & vbNewline
						i = i + 1
		            		loop 
					Response.Write	"                </select></td>" & vbNewline & _
							"              </tr>" & vbNewline
				 	forumordercount = forumordercount + 1
					rsForum.MoveNext
				loop
			end if
			catordercount = catordercount + 1	
			rs.MoveNext
		loop
		rsForum.close
		set rsForum = nothing 
	Response.Write	"              <tr valign=""top"">" & _
			"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center"">" & vbNewline & _
			"                <input type=""submit"" value=""Submit Order"" id=""submit1"" name=""submit1"">&nbsp;<input type=""reset"" value=""Reset Old Values"" & id=""reset1"" name=""reset1""></td>" & vbNewline & _
			"              </tr>" & vbNewline
	end if 
	Response.Write	"            </table>" & vbNewline & _
			"          </td>" & vbNewline & _
			"        </tr>" & vbNewline & _
			"      </table>" & vbNewline & _
			"      </form>" & vbNewline
	rs.close
	set rs = nothing 
end if
WriteFooterShort
Response.End
%>