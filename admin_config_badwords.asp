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
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login_short.asp?target=" & scriptname(ubound(scriptname))
end if
strRqMethod = trim(chkString(Request.QueryString("method"),"SQLString"))
intBadwordID = trim(chkString(Request.QueryString("B_ID"),"SQLString"))

if intBadwordID <> "" then
	if isNumeric(intBadwordID) <> True then intBadwordID = "0"
end if

strPageSize = 10

mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)

Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"    <!-----" & vbNewLine & _
		"    function jumpToPage(s) {location.href = s.options[s.selectedIndex].value}" & vbNewLine & _
		"    // -->" & vbNewLine & _
		"    </script>" & vbNewLine

Select Case strRqMethod
	Case "Add"
		if Request.Form("Method_Type") = "Write_Configuration" then 
			Err_Msg = ""

			txtBadword = chkBString(Request.Form("strBadword"),"SQLString")
			txtReplace = chkBString(Request.Form("strReplace"),"SQLString")
	
			if txtBadword = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Badword.</li>"
			end if

			if txtBadword = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Badword.</li>"
			end if

			if (Instr(txtBadword, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the Badword.</li>"
			end if

			if txtReplace = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Replacement word for the Badword.</li>"
			end if

			if txtReplace = "" then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Replacement word for the Badword.</li>"
			end if

			if (Instr(txtReplace, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the Replacement word.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "INSERT INTO " & strFilterTablePrefix & "BADWORDS ("
				strSql = strSql & "B_BADWORD"
				strSql = strSql & ", B_REPLACE"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & txtBadword & "'"
				strSql = strSql & ", '" & txtReplace & "'"
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Application.Lock
				Application(strCookieURL & "STRBADWORDWORDS") = ""
				Application(strCookieURL & "STRBADWORDREPLACE") = ""
				Application.UnLock

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Badword Added!</font></p>" & vbNewLine & _
						"      <meta http-equiv=""Refresh"" content=""1; URL=admin_config_badwords.asp"">" & vbNewLine & _
						"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""admin_config_badwords.asp"">Back To Badword Filter Configuration</font></a></p>" & vbNewLine
			else
				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>There Was A Problem With Your Details</font></p>" & vbNewLine & _
						"      <table align=""center"" border=""0"">" & vbNewLine & _
						"        <tr>" & vbNewLine & _
						"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHiLiteFontColor & """><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
						"        </tr>" & vbNewLine & _
						"      </table>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></font></p>" & vbNewLine
			end if
		end if
	Case "Delete"
		if Request.Form("Method_Type") = "Delete_Badword" then
			'## Forum_SQL - Delete badword from Badwords table
			strSql = "DELETE FROM " & strFilterTablePrefix & "BADWORDS "
			strSql = strSql & " WHERE B_ID = " & Request.Form("B_ID")

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			Application.Lock
			Application(strCookieURL & "STRBADWORDWORDS") = ""
			Application(strCookieURL & "STRBADWORDREPLACE") = ""
			Application.UnLock

			Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Badword Deleted!</b></font></p>" & vbNewLine & _
					"      <meta http-equiv=""Refresh"" content=""1; URL=admin_config_badwords.asp"">" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_badwords.asp"">Back To Badword Filter Configuration</font></a></p>" & vbNewLine
		else
			Response.Write	"      <form action=""admin_config_badwords.asp?method=Delete"" method=""post"" id=""UpdateBWord"" name=""UpdateBWord"">" & vbNewLine & _
					"      <input type=""hidden"" name=""Method_Type"" value=""Delete_Badword"">" & vbNewLine & _
					"      <input type=""hidden"" name=""B_ID"" value=""" & intBadwordID & """>" & vbNewLine & _
					"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Are you sure?</b></font></p>" & vbNewLine & _
					"      <p align=""center""><input type=""submit"" class=""button"" value=""Yes"" id=""submit1"" name=""submit1"">&nbsp;<input type=""button"" class=""button"" value="" No "" onClick=""history.go(-1);""></p>" & vbNewLine & _
					"      </form>" & vbNewLine
		end if
	Case "Edit"
		if Request.Form("Method_Type") = "Write_Configuration" then
			txtBadword = chkBString(Request.Form("strBadword"),"SQLString")
			txtReplace = chkBString(Request.Form("strReplace"),"SQLString")

			if txtBadword = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Badword.</li>"
			end if

			if (Instr(txtBadword, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the Badword.</li>"
			end if

			if txtReplace = " " then 
				Err_Msg = Err_Msg & "<li>You Must Enter a Replacement word for the Badword.</li>"
			end if

			if (Instr(txtReplace, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the Replacement word.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strFilterTablePrefix & "BADWORDS "
				strSql = strSql & " SET B_BADWORD = '" & txtBadword & "'"
				strSql = strSql & ",    B_REPLACE = '" & txtReplace & "'"
				strSql = strSql & " WHERE B_ID = " & Request.Form("B_ID")

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Application.Lock
				Application(strCookieURL & "STRBADWORDWORDS") = ""
				Application(strCookieURL & "STRBADWORDREPLACE") = ""
				Application.UnLock

				Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Badword Filter Updated!</font></p>" & vbNewLine & _
						"      <meta http-equiv=""Refresh"" content=""1; URL=admin_config_badwords.asp"">" & vbNewLine & _
						"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Congratulations!</font></p>" & vbNewLine & _
						"      " & strParagraphFormat1 & "<a href=""admin_config_badwords.asp"">Back To Badword Filter Configuration</font></a></p>" & vbNewLine
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
			'## Forum_SQL - Get Badword/Replacement word from DB
			strSql = "SELECT B_ID, B_BADWORD, B_REPLACE "
			strSql = strSql & " FROM " & strFilterTablePrefix & "BADWORDS "
			strSql = strSql & " WHERE B_ID = " & intBadwordID

			set rs = my_Conn.Execute (strSql)

		        TxtBadword = rs("B_BADWORD")
		        TxtReplace = rs("B_REPLACE")
                        intB_ID = rs("B_ID")

                        rs.close
                        set rs = nothing

			Response.Write	"      <form action=""admin_config_badwords.asp?method=Edit"" method=""post"" id=""UpdateBWord"" name=""UpdateBWord"">" & vbNewLine & _
					"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
					"      <input type=""hidden"" name=""B_ID"" value=""" & intB_ID & """>" & vbNewLine & _
					"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
					"        <tr>" & vbNewLine & _
					"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
					"            <table border=""0"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td align=""center"" bgcolor=""" & strCategoryCellColor & """ colspan=""2""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Edit Badword</b></font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Badword</b></font></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>Replacement</b></font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><input maxLength=""50"" name=""strBadword"" value=""" & TxtBadword & """ size=""12"" tabindex=""1""></td>" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><input maxLength=""50"" name=""strReplace"" value=""" & TxtReplace & """ size=""12"" tabindex=""2""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"              <tr valign=""middle"">" & vbNewLine & _
					"                <td bgColor=""" & strPopUpTableColor & """ colspan=""2"" align=""center""><input type=""submit"" class=""button"" value=""Update"" id=""submit1"" name=""submit1"" tabindex=""3""> <input type=""reset"" class=""button"" value=""Reset"" id=""reset1"" name=""reset1""></td>" & vbNewLine & _
					"              </tr>" & vbNewLine & _
					"            </table>" & vbNewLine & _
					"          </td>" & vbNewLine & _
					"        </tr>" & vbNewLine & _
					"      </table>" & vbNewLine & _
					"      </form>" & vbNewLine & _
					"      " & strParagraphFormat1 & "<a href=""admin_config_badwords.asp"">Back To Badword Filter Configuration</font></a></p>" & vbNewLine
		end if
	Case Else
		'## Forum_SQL - Get Badwords from DB
		strSql = "SELECT B_ID, B_BADWORD, B_REPLACE " 
		strSql2 = " FROM " & strFilterTablePrefix & "BADWORDS "
		strSql3 = " ORDER BY B_BADWORD ASC "

		if strDBType = "mysql" then 'MySql specific code
			if mypage > 1 then 
				OffSet = cLng((mypage - 1) * strPageSize)
				strSql4 = " LIMIT " & OffSet & ", " & strPageSize & " "
			end if

			'## Forum_SQL - Get the total pagecount 
			strSql1 = "SELECT COUNT(B_ID) AS PAGECOUNT "

			set rsCount = my_Conn.Execute(strSql1 & strSql2)
			iPageTotal = rsCount(0).value
			rsCount.close
			set rsCount = nothing

			If iPageTotal > 0 then
				maxpages = (iPageTotal \ strPageSize )
				if iPageTotal mod strPageSize <> 0 then
					maxpages = maxpages + 1
				end if
				if iPageTotal < (strPageSize + 1) then
					intGetRows = iPageTotal
				elseif (mypage * strPageSize) > iPageTotal then
					intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
				else
					intGetRows = strPageSize
				end if
			else
				iPageTotal = 0
				maxpages = 0
			end if 

			if iPageTotal > 0 then
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					arrBadwordData = rs.GetRows(intGetRows)
					iBadwordCount = UBound(arrBadwordData, 2)
				rs.close
				set rs = nothing
			else
				iBadwordCount = ""
			end if
 
		else 'end MySql specific code

			set rs = Server.CreateObject("ADODB.Recordset")
			rs.cachesize = strPageSize
			rs.open strSql & strSql2 & strSql3, my_Conn, adOpenStatic
				If not (rs.EOF or rs.BOF) then
					rs.movefirst
					rs.pagesize = strPageSize
					rs.absolutepage = mypage '**
					maxpages = cLng(rs.pagecount)
					arrBadwordData = rs.GetRows(strPageSize)
					iBadwordCount = UBound(arrBadwordData, 2)
				else
					iBadwordCount = ""
				end if
			rs.Close
			set rs = nothing
		end if

		Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Bad Word Filter Configuration</b></font></p>" & vbNewLine
		Response.Write	"      <form action=""admin_config_badwords.asp?method=Add"" method=""post"" id=""Add"" name=""Add"">" & vbNewLine & _
				"      <input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & vbNewLine & _
				"      <table border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
				"            <table width=""100%"" align=""center"" border=""0"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td align=""center"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Badword</font></b></td>" & vbNewLine & _
				"                <td align=""center"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>Replacement</font></b></td>" & vbNewLine
		if maxpages > 1 then
			Call DropDownPaging()
		else
			Response.Write	"                <td align=""center"" bgcolor=""" & strCategoryCellColor & """><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """>&nbsp;</font></b></td>" & vbNewLine
		end if
		Response.Write	"              </tr>" & vbNewLine

		if iBadwordCount = "" then  '## No Badwords found in DB
			Response.Write	"              <tr>" & vbNewLine & _
        				"                <td bgcolor=""" & strForumFirstCellColor & """ colspan=""3""><font face=""" & strDefaultFontFace & """ color=""" & strForumFontColor & """ size=""" & strDefaultFontSize & """ valign=""top""><b>No Badwords Found</b></font></td>" & vbNewLine & _
					"              </tr>" & vbNewLine
		else
			bB_ID = 0
			bB_BADWORD = 1
			bB_REPLACE = 2

			rec = 1
			intI = 0

			for iBadword = 0 to iBadwordCount
				if (rec = strPageSize + 1) then exit for

				Badword_ID = arrBadwordData(bB_ID, iBadword)
				Badword_Badword = arrBadwordData(bB_BADWORD, iBadword)
				Badword_Replace = arrBadwordData(bB_REPLACE, iBadword)

				if intI = 1 then 
					CColor = strAltForumCellColor
				else
					CColor = strForumCellColor
				end if

				Response.Write	"              <tr>" & vbNewLine & _
						"                <td bgcolor=""" & CColor & """ valign=""middle"" align=""center""><font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & Badword_Badword & "</font></td>" & vbNewLine & _
						"                <td bgcolor=""" & CColor & """ valign=""middle"" align=""center""><font color=""" & strForumFontColor & """ face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & Badword_Replace & "</font></td>" & vbNewLine & _
						"                <td bgcolor=""" & CColor & """ valign=""middle"" align=""center"" nowrap><a href=""admin_config_badwords.asp?method=Edit&B_ID=" & Badword_ID & """>" & getCurrentIcon(strIconPencil,"Edit Badword","hspace=""0""") & "</a>&nbsp;<a href=""admin_config_badwords.asp?method=Delete&B_ID=" & Badword_ID & """>" & getCurrentIcon(strIconTrashcan,"Delete Badword","hspace=""0""") & "</a></td>" & vbNewLine & _
						"              </tr>" & vbNewLine
				rec = rec + 1
				intI = intI + 1
				if intI = 2 then
					intI = 0
				end if
			next
		end if

		Response.Write	"              <tr valign=""middle"">" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><input maxLength=""50"" name=""strBadword"" value=""" & TxtBadword & """ tabindex=""1"" size=""10""></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""center""><input maxLength=""50"" name=""strReplace"" value=""" & TxtReplace & """ tabindex=""2"" size=""10""></td>" & vbNewLine & _
				"                <td bgcolor=""" & strPopUpTableColor & """ valign=""middle"" align=""center"" nowrap><input class=""button"" value=""Add"" type=""submit"" tabindex=""3""></a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      </form>" & vbNewLine
End Select
WriteFooterShort
Response.End

sub DropDownPaging()
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		Response.Write	"                <td valign=""middle"" bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strCategoryFontColor & """>" & vbNewLine & _
				"                <b>Page</b>&nbsp;<select style=""font-size:9px"" name=""whichpage"" size=""1"" onchange=""jumpToPage(this)"">" & vbNewLine
		for counter = 1 to maxpages
			ref = "admin_config_badwords.asp?whichpage=" & counter 
			if counter <> cLng(pge) then
				Response.Write	"                	<option value=""" & ref & """>" & counter & "</option>" & vbNewLine
			else
				Response.Write	"                	<option value=""" & ref & """ selected>" & counter & "</option>" & vbNewLine
			end if
		next
		Response.Write	"                </select>&nbsp;<b>of " & maxpages & "</b></font></td>" & vbNewLine
	end if
end sub 

function chkBString(fString,fField_Type) '## Types - SQLString
	if fString = "" then
		fString = " "
	end if
	Select Case fField_Type
		Case "SQLString"
			fString = Replace(fString, "'", "''")
			if strDBType = "mysql" then
				fString = Replace(fString, "\0", "\\0")
				fString = Replace(fString, "\'", "\\'")
				fString = Replace(fString, "\""", "\\""")
				fString = Replace(fString, "\b", "\\b")
				fString = Replace(fString, "\n", "\\n")
				fString = Replace(fString, "\r", "\\r")
				fString = Replace(fString, "\t", "\\t")
				fString = Replace(fString, "\z", "\\z")
				fString = Replace(fString, "\%", "\\%")
				fString = Replace(fString, "\_", "\\_")
			end if
			fString = HTMLBEncode(fString)
			chkBString = fString
			exit function
	End Select
	chkBString = fString
end function

function HTMLBEncode(fString)
	if fString = "" or IsNull(fString) then fString = " "
	fString = replace(fString, ">", "&gt;")
	fString = replace(fString, "<", "&lt;")
	HTMLBEncode = fString
end function
%>