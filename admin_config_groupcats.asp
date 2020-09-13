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
	if Request.QueryString <> "" then
		Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname)) & "?" & Request.QueryString
	else
		Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
	end if
end if
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;<a href=""admin_config_groupcats.asp"">Group&nbsp;Categories&nbsp;Configuration</a><br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine

Response.Write	"    <script language=""JavaScript"" type=""text/javascript"" src=""selectbox.js""></script>" & vbNewLine

strRqMethod = Request.QueryString("method")

Select Case strRqMethod
	Case "Add"
		if Request.Form("Method_Type") = "Write_Configuration" then 
			Err_Msg = ""

			txtGroupName = chkString(Request.Form("strGroupName"),"SQLString")
			txtGroupDescription = chkString(Request.Form("strGroupDescription"),"message")
			txtGroupIcon = chkString(Request.Form("strGroupIcon"),"SQLString")
			txtGroupTitleImage = chkString(Request.Form("strGroupTitleImage"),"SQLString")

			if trim(txtGroupName) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Name for your New Group.</li>"
			end if

			if trim(txtGroupDescription) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Description for your New Group.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES ("
				strSql = strSql & "GROUP_NAME"
				strSql = strSql & ", GROUP_DESCRIPTION"
				strSql = strSql & ", GROUP_ICON"
				strSql = strSql & ", GROUP_IMAGE"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & txtGroupName & "'"
				strSql = strSql & ", '" & txtGroupDescription & "'"
				strSql = strSql & ", '" & txtGroupIcon & "'"
				strSql = strSql & ", '" & txtGroupTitleImage & "'"
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				set rsCount = my_Conn.execute("SELECT MAX(GROUP_ID) AS maxGroupID FROM " & strTablePrefix & "GROUP_NAMES ")
				newGroupCategories rsCount("maxGroupId")
				set rsCount = nothing

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>New Group Added!</font></p>" & vbNewLine & _
						"      <meta http-equiv=""Refresh"" content=""2; URL=admin_config_groupcats.asp"">" & vbNewLine & _
						"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
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
			Response.Write	"      <form action=""admin_config_groupcats.asp?method=Add"" method=""post"" id=""Add"" name=""Add"">" & vbNewLine & _
					"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
					"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Create A New Category Group</font></b></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>New Group Name:</b>&nbsp;</font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""50"" name=""strGroupName"" value="""" tabindex=""1"" size=""46""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""top"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>New Group Description:</b>&nbsp;</font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><textarea maxLength=""255"" rows=""5"" cols=""35"" name=""strGroupDescription"" tabindex=""2""></textarea></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>New Group Icon:</b>&nbsp;</font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""255"" name=""strGroupIcon"" value="""" tabindex=""3"" size=""46""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>New Group Title Image:</b>&nbsp;</font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""255"" name=""strGroupTitleImage"" value="""" tabindex=""4"" size=""46""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr>" & vbNewLine & _
				      	"                <td bgColor=""" & strPopUpTableColor & """ noWrap vAlign=""top"" align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Categories:</b>&nbsp;</font></td>" & vbNewLine
			strSql = "SELECT CAT_ID, CAT_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
			strSql = strSql & " ORDER BY CAT_NAME ASC "

			set rsCats = Server.CreateObject("ADODB.Recordset")
			rsCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rsCats.EOF then
				recCatCnt = ""
			else
				allCatData = rsCats.GetRows(adGetRowsRest)
				recCatCnt = UBound(allCatData,2)
				cCAT_ID = 0
				cCAT_NAME = 1
			end if

			rsCats.close
			set rsCats = nothing

			SelectSize = 6
			Response.Write	"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
					"                  <table>" & vbNewLine & _
					"                    <tr>" & vbNewLine & _
					"                      <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Available</b></font><br />" & vbNewLine & _
					"                      <select name=""GroupCatCombo"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & vbNewLine
			'## Pick from list
			if recCatCnt <> "" then
				for iCat = 0 to recCatCnt
					CategoryCatID = allCatData(cCAT_ID,iCat)
					CategoryCatName = allCatData(cCAT_NAME,iCat)
					Response.Write 	"                      		<option value=""" & CategoryCatID & """>" & ChkString(CategoryCatName,"display") & "</option>" & vbNewline
				next
			end if
			Response.Write	"                      </select>" & vbNewLine & _
					"                      </td>" & vbNewLine & _
					"                      <td width=""15"" align=""center"" valign=""middle""><br />" & vbNewLine & _
					"                      <a href=""javascript:moveAllOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemAll,"","") & "</a>" & vbNewLine & _
					"                      <a href=""javascript:moveSelectedOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemove,"","") & "</a>" & vbNewLine & _
					"                      <a href=""javascript:moveSelectedOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAdd,"","") & "</a>" & vbNewLine & _
					"                      <a href=""javascript:moveAllOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAddAll,"","") & "</a>" & vbNewLine & _
					"                      </td>" & vbNewLine & _
					"                      <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Selected</b></font><br />" & vbNewLine & _
					"                      <select name=""GroupCat"" size=""" & SelectSize & """ multiple tabindex=""15"" onDblClick=""moveSelectedOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & vbNewLine & _
					"                      </select>" & vbNewLine & _
					"                      </td>" & vbNewLine & _
					"                    </tr>" & vbNewLine & _
					"                  </table>" & vbNewLine & _
					"                </td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """ valign=""middle"" align=""center"" colspan=""2"" nowrap><input class=""button"" value=""  Add  "" type=""submit"" tabindex=""5"" onclick=""selectAllOptions(document.Add.GroupCat);"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset"" tabindex=""6""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"            </table>" & vbNewLine & _
					"          </td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      </form>" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
		end if
	Case "Delete"
		if Request.Form("Method_Type") = "Delete_Category" then
			'## Forum_SQL
			strSql = "DELETE FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			strSql = "DELETE FROM " & strTablePrefix & "GROUPS "
			strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Category Group Deleted!</b></font></p>" & vbNewLine & _
					"      <meta http-equiv=""Refresh"" content=""1; URL=admin_config_groupcats.asp"">" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
		else
			'## Forum_SQL
			strSql = "SELECT GROUP_ID, GROUP_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID <> 1 "
			strSql = strSql & " AND GROUP_ID <> 2 "
			strSql = strSql & " ORDER BY GROUP_NAME ASC "

			Set rsgroups = Server.CreateObject("ADODB.Recordset")
			rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			If rsgroups.EOF then
				recGroupCount = ""
			Else
				allGroupData = rsgroups.GetRows(adGetRowsRest)
				recGroupCount = UBound(allGroupData, 2)
			End if

			rsgroups.Close
			Set rsgroups = Nothing

			Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
					"    <!--" & vbNewLine & _
					"    function confirmDelete(){" & vbNewLine & _
					"    	var where_to= confirm(""Do you really want to Delete this Group Category?"", ""Yes"", ""No"");" & vbNewLine & _
					"       if (where_to)" & vbNewLine & _
					"       	return true;" & vbNewLine & _
					"       else" & vbNewLine & _
					"       	return false;" & vbNewLine & _
					"    }" & vbNewLine & _
					"    //-->" & vbNewLine & _
					"    </script>" & vbNewLine

			Response.Write	"      <form action=""admin_config_groupcats.asp?method=Delete"" method=""post"" id=""Add"" name=""Add"">" & vbNewLine & _
					"      <input type=""hidden"" name=""Method_Type"" value=""Delete_Category"">" & vbNewLine & _
					"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
					"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Delete Group Categories</font></b></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine
			if recGroupCount <> "" then
				Response.Write	"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Choose Group To Delete:</b>&nbsp;</font></td>" & vbNewLine & _
						"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
						"                <select name=""GroupID"" size=""1"">" & vbNewLine
				for iGroup = 0 to recGroupCount
					Response.Write	"                	<option value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0,iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & vbNewLine
				next
				Response.Write	"                </select>" & vbNewLine & _
						"                </font></td>" & vbNewLine & _
						"              </tr>" & vbNewLine & _
						"              <tr valign=""middle"">" & vbNewLine & _
						"                <td bgcolor=""" & strPopUpTableColor & """ valign=""middle"" align=""center"" colspan=""2"" nowrap><input class=""button"" value="" Delete "" type=""submit"" onClick=""return confirmDelete()"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset""></td>" & vbNewLine
			else
				Response.Write	"                <td bgColor=""" & strPopUpTableColor & """ align=""center"" colspan=""2"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b><i>No Groups Available To Delete</i></b>&nbsp;</font></td>" & vbNewLine
			end if
			Response.Write	"              </tr>" & vbNewLine & _
					"            </table>" & vbNewLine & _
					"          </td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      </form>" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
		end if
	Case "Edit"
		if Request.Form("Method_Type") = "Write_Configuration" then
			txtGroupName = chkString(Request.Form("strGroupName"),"SQLString")
			txtGroupDescription = chkString(Request.Form("strGroupDescription"),"message")
			txtGroupIcon = chkString(Request.Form("strGroupIcon"),"SQLString")
			txtGroupTitleImage = chkString(Request.Form("strGroupTitleImage"),"SQLString")

			if trim(txtGroupName) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Name for your New Group.</li>"
			end if

			if trim(txtGroupDescription) = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Description for your New Group.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strTablePrefix & "GROUP_NAMES "
				strSql = strSql & " SET GROUP_NAME = '" & txtGroupName & "'"
				strSql = strSql & ",    GROUP_DESCRIPTION = '" & txtGroupDescription & "'"
				strSql = strSql & ",    GROUP_ICON = '" & txtGroupIcon & "'"
				strSql = strSql & ",    GROUP_IMAGE = '" & txtGroupTitleImage & "'"
				strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GROUP_ID"))

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				updateGroupCategories(Request.Form("GROUP_ID"))

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Category Group Updated!</font></p>" & vbNewLine & _
						"      <meta http-equiv=""Refresh"" content=""2; URL=admin_config_groupcats.asp"">" & vbNewLine & _
						"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
						"      <table align=""center"" border=""0"">" & vbNewLine & _
						"        <tr>" & vbNewLine & _
						"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
						"        </tr>" & vbNewLine & _
						"      </table>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">Go Back To Correct The Problem</a></font></p>" & vbNewLine
			end if
		elseif Request.Form("Method_Type") = "Edit_Category" then 
			if Request.Form("GroupID") <> "" then
				'## Forum_SQL
				strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_DESCRIPTION, GROUP_ICON, GROUP_IMAGE  "
				strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
				strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

				set rsGroups = Server.CreateObject("ADODB.Recordset")
				rsGroups.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

				if rsGroups.EOF then
					recGroupCnt = ""
				else
					allGroupData = rsGroups.GetRows(adGetRowsRest)
					recGroupCnt = UBound(allGroupData,2)
					gGROUP_ID = 0
					gGROUP_NAME = 1
					gGROUP_DESCRIPTION = 2
					gGROUP_ICON = 3
					gGROUP_IMAGE = 4
				end if

				rsGroups.close
				set rsGroups = nothing

				if recGroupCnt <> "" then
					txtGroupID = allGroupData(gGROUP_ID,0)
					txtGroupName = allGroupData(gGROUP_NAME,0)
					txtGroupDescription = allGroupData(gGROUP_DESCRIPTION,0)
					txtGroupIcon = allGroupData(gGROUP_ICON,0)
					txtGroupTitleImage = allGroupData(gGROUP_IMAGE,0)

					Response.Write	"      <form action=""admin_config_groupcats.asp?method=Edit"" method=""post"" id=""Edit"" name=""Edit"">" & vbNewLine & _
							"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
							"      <input type=""hidden"" name=""GROUP_ID"" value=""" & txtGroupID & """>" & vbNewLine & _
							"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
							"        <tr>" & vbNewLine & _
							"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
							"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
							"              <tr>" & vbNewLine & _
							"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Edit Existing Category Group</font></b></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr valign=""middle"">" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Group Name:</b>&nbsp;</font></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""50"" name=""strGroupName"" value=""" & txtGroupName & """ tabindex=""1"" size=""46""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr valign=""top"">" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Group Description:</b>&nbsp;</font></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><textarea rows=""5"" cols=""35"" name=""strGroupDescription"" maxLength=""255"" tabindex=""2"">" & txtGroupDescription & "</textarea></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr valign=""middle"">" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Group Icon:</b>&nbsp;</font></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""255"" name=""strGroupIcon"" value=""" & txtGroupIcon & """ tabindex=""3"" size=""46""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr valign=""middle"">" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Group Title Image:</b>&nbsp;</font></td>" & vbNewLine & _
							"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><input maxLength=""255"" name=""strGroupTitleImage"" value=""" & txtGroupTitleImage & """ tabindex=""4"" size=""46""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr>" & vbNewLine & _
						      	"                <td bgColor=""" & strPopUpTableColor & """ noWrap vAlign=""top"" align=""right""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Categories:</b>&nbsp;</font></td>" & vbNewLine
					strSql = "SELECT CAT_ID, CAT_NAME "
					strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
					strSql = strSql & " ORDER BY CAT_NAME ASC "

					set rsCats = Server.CreateObject("ADODB.Recordset")
					rsCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

					if rsCats.EOF then
						recCatCnt = ""
					else
						allCatData = rsCats.GetRows(adGetRowsRest)
						recCatCnt = UBound(allCatData,2)
						cCAT_ID = 0
						cCAT_NAME = 1
					end if

					rsCats.close
					set rsCats = nothing

					tmpStrUserList  = ""

					strSql = "SELECT GROUP_CATID "
					strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
					strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

					set rsGroupCats = Server.CreateObject("ADODB.Recordset")
					rsGroupCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

					if rsGroupCats.EOF then
						recGroupCatCnt = ""
					else
						allGroupCatData = rsGroupCats.GetRows(adGetRowsRest)
						recGroupCatCnt = UBound(allGroupCatData,2)
						gGROUP_CATID = 0
					end if

					rsGroupCats.close
					set rsGroupCats = nothing

					if recGroupCatCnt <> "" then
						for iGroupCats = 0 to recGroupCatCnt
							GroupCatID = allGroupCatData(gGROUP_CATID,iGroupCats)
							if tmpStrUserList = "" then
								tmpStrUserList = GroupCatID
							else
								tmpStrUserList = tmpStrUserList & "," & GroupCatID
							end if
						next
					end if
					SelectSize = 6
					Response.Write	"                <td bgColor=""" & strPopUpTableColor & """>" & vbNewLine & _
							"                  <table>" & vbNewLine & _
							"                    <tr>" & vbNewLine & _
							"                      <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Available</b></font><br />" & vbNewLine & _
							"                      <select name=""GroupCatCombo"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & vbNewLine
					'## Pick from list
					if recCatCnt <> "" then
						for iCat = 0 to recCatCnt
							CategoryCatID = allCatData(cCAT_ID,iCat)
							CategoryCatName = allCatData(cCAT_NAME,iCat)
							if not(Instr("," & tmpStrUserList & "," , "," & CategoryCatID & ",") > 0) then
								Response.Write 	"                      		<option value=""" & CategoryCatID & """>" & ChkString(CategoryCatName,"display") & "</option>" & vbNewline
							end if
						next
					end if
					Response.Write	"                      </select>" & vbNewLine & _
							"                      </td>" & vbNewLine & _
							"                      <td width=""15"" align=""center"" valign=""middle""><br />" & vbNewLine & _
							"                      <a href=""javascript:moveAllOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemAll,"","") & "</a>" & vbNewLine & _
							"                      <a href=""javascript:moveSelectedOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemove,"","") & "</a>" & vbNewLine & _
							"                      <a href=""javascript:moveSelectedOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAdd,"","") & "</a>" & vbNewLine & _
							"                      <a href=""javascript:moveAllOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAddAll,"","") & "</a>" & vbNewLine & _
							"                      </td>" & vbNewLine & _
							"                      <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Selected</b></font><br />" & vbNewLine & _
							"                      <select name=""GroupCat"" size=""" & SelectSize & """ multiple tabindex=""15"" onDblClick=""moveSelectedOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & vbNewLine
					if recGroupCatCnt <> "" then
						for iGroupCats = 0 to recGroupCatCnt
							GroupCatID = allGroupCatData(gGROUP_CATID,iGroupCats)
							if GroupCatID <> "" then
								Response.Write 	"                      		<option value=""" & GroupCatID & """>" & ChkString(getCategoryName(GroupCatID),"display") & "</option>" & vbNewline
							end if
						next
					end if
					Response.Write	"                      </select>" & vbNewLine & _
							"                      </td>" & vbNewLine & _
							"                    </tr>" & vbNewLine & _
							"                  </table>" & vbNewLine & _
							"                </td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"              <tr valign=""middle"">" & vbNewLine & _
							"                <td bgcolor=""" & strPopUpTableColor & """ valign=""middle"" align=""center"" colspan=""2"" nowrap><input class=""button"" value=""Submit"" type=""submit"" tabindex=""5"" onclick=""selectAllOptions(document.Edit.GroupCat);"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset"" tabindex=""6""></td>" & vbNewLine & _
							"              </tr>" & vbNewLine & _
							"            </table>" & vbNewLine & _
							"          </td>" & vbNewLine & _
							"        </tr>" & vbNewLine & _
							"      </table>" & vbNewLine & _
							"      </form>" & vbNewLine & _
							"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
				else
					Response.Write	"      " & strParagraphFormat1 & "Invalid Group ID</font></p>" & vbNewLine & _
							"      <br />" & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
				end if
			else
				Response.Write	"      " & strParagraphFormat1 & "Invalid Group ID</font></p>" & vbNewLine & _
						"      <br />" & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">Go back to correct the problem.</a></font></p>" & vbNewLine
			end if
		else
			'## Forum_SQL
			strSql = "SELECT GROUP_ID, GROUP_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID <> 1 "
			strSql = strSql & " ORDER BY GROUP_NAME ASC "

			Set rsgroups = Server.CreateObject("ADODB.Recordset")
			rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			If rsgroups.EOF then
				recGroupCount = ""
			Else
				allGroupData = rsgroups.GetRows(adGetRowsRest)
				recGroupCount = UBound(allGroupData, 2)
			End if

			rsgroups.Close
			Set rsgroups = Nothing

			Response.Write	"      <form action=""admin_config_groupcats.asp?method=Edit"" method=""post"" id=""Add"" name=""Add"">" & vbNewLine & _
					"      <input type=""hidden"" name=""Method_Type"" value=""Edit_Category"">" & vbNewLine & _
					"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
					"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
					"              <tr>" & vbNewLine & _
					"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Edit Group Categories</font></b></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>&nbsp;<b>Choose Group To Edit:</b>&nbsp;</font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
					"                <select name=""GroupID"" size=""1"">" & vbNewLine
			if recGroupCount <> "" then
				for iGroup = 0 to recGroupCount
					if allGroupData(0, iGroup) = 2 then
						Response.Write	"                	<option label=""" & chkString(allGroupData(1, iGroup),"display") & """ value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0, iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & vbNewLine
						exit for
					end if
				next
				for iGroup = 0 to recGroupCount
					if allGroupData(0, iGroup) <> 2 then
						first = 0
						if first = 0 then
							Response.Write	"                	<option value="""">----------------------------</option>" & vbNewLine
							first = 1
						end if
						Response.Write	"                	<option value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0, iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & vbNewLine
					end if
				next
			end if
			Response.Write	"                </select>" & vbNewLine & _
					"                </font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgcolor=""" & strPopUpTableColor & """ valign=""middle"" align=""center"" colspan=""2"" nowrap><input class=""button"" value=""  Edit  "" type=""submit"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"            </table>" & vbNewLine & _
					"          </td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      </form>" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</font></a></p>" & vbNewLine
		end if
	Case Else
		Response.Write	"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
				"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td align=""left"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Group Categories Configuration</font></b></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr valign=""middle"">" & vbNewLine & _
				"                <td bgColor=""" & strForumCellColor & """ align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><UL>" & vbNewLine & _
				"                <LI><span class=""spnMessageText""><a href=""admin_config_groupcats.asp?method=Add"">Create A New Category Group</a></span></LI>" & vbNewLine & _
				"                <LI><span class=""spnMessageText""><a href=""admin_config_groupcats.asp?method=Delete"">Delete A Category Group</a></span></LI>" & vbNewLine & _
				"                <LI><span class=""spnMessageText""><a href=""admin_config_groupcats.asp?method=Edit"">Edit an Existing Category Group</a></span>&nbsp;&nbsp;&nbsp;&nbsp;</LI>" & vbNewLine & _
				"                </UL></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      </form>" & vbNewLine
End Select
Response.Write	"      <br />" & vbNewLine
WriteFooter
Response.End

sub newGroupCategories(fGroupID)
	if Request.Form("GroupCat") = "" then
		exit Sub
	end if
	Cats = split(Request.Form("GroupCat"),",")
	for count = Lbound(Cats) to Ubound(Cats)
		strSql = "INSERT INTO " & strTablePrefix & "GROUPS ("
		strSql = strSql & " GROUP_ID, GROUP_CATID) VALUES ( "& fGroupID & ", " & Cats(count) & ")"
		my_conn.execute (strSql),,adCmdText + adExecuteNoRecords
	next
end sub

sub updateGroupCategories(fGroupID)
	my_Conn.execute ("DELETE FROM " & strTablePrefix & "GROUPS WHERE GROUP_ID = " & fGroupId),,adCmdText + adExecuteNoRecords
	newGroupCategories(fGroupID)
end sub

Function getCategoryName(fCat_ID)
	set rsCatName = my_Conn.execute("SELECT CAT_NAME FROM " & strTablePrefix & "CATEGORY WHERE CAT_ID = " & fCat_ID)
	getCategoryName = rsCatName("CAT_NAME")
	set rsCatName = nothing
end function
%>
