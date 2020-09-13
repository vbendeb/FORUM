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
Response.Write	"    <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"    <!--" & vbNewLine & _
		"    function ChangePage(fnum){" & vbNewLine & _
		"    	if (fnum == 1) {" & vbNewLine & _
		"    		document.PageNum1.submit();" & vbNewLine & _
		"    	}" & vbNewLine & _
		"    	else {" & vbNewLine & _
		"    		document.PageNum2.submit();" & vbNewLine & _
		"    	}" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function appr_all(){" & vbNewLine & _
		"    	var where_to= confirm(""Do you really want to Approve all Pending Members?"");" & vbNewLine & _
		"       if (where_to== true) {" & vbNewLine & _
		"       	window.location=""admin_accounts_pending.asp?id=-1&action=approve"";" & vbNewLine & _
		"       }" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function appr_selected(){" & vbNewLine & _
		"    	var where_to= confirm(""Do you really want to Approve the Selected Pending Members?"");" & vbNewLine & _
		"       if (where_to== true) {" & vbNewLine & _
		"		document.delMembers.action.value = 'approve';" & vbNewLine & _
		"    		document.delMembers.submit();" & vbNewLine & _
		"       }" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function del_all(){" & vbNewLine & _
		"    	var where_to= confirm(""Do you really want to Delete all Pending Members?"");" & vbNewLine & _
		"       if (where_to== true) {" & vbNewLine & _
		"       	window.location=""admin_accounts_pending.asp?id=-1&action=delete"";" & vbNewLine & _
		"       }" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function del_selected(){" & vbNewLine & _
		"    	var where_to= confirm(""Do you really want to Delete the Selected Pending Members?"");" & vbNewLine & _
		"       if (where_to== true) {" & vbNewLine & _
		"		document.delMembers.action.value = 'delete';" & vbNewLine & _
		"       	document.delMembers.submit();" & vbNewLine & _
		"       }" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function Toggle(field)" & vbNewLine & _
		"    {" & vbNewLine & _
		"	if (field.checked) {" & vbNewLine & _
		"	    document.delMembers.toggleAll.checked = AllChecked();" & vbNewLine & _
		"	}" & vbNewLine & _
		"	else {" & vbNewLine & _
		"	    document.delMembers.toggleAll.checked = false;" & vbNewLine & _
		"	}" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function ToggleAll(field)" & vbNewLine & _
		"    {" & vbNewLine & _
		"	if (field.checked) {" & vbNewLine & _
		"	    CheckAll();" & vbNewLine & _
		"	}" & vbNewLine & _
		"	else {" & vbNewLine & _
		"	    ClearAll();" & vbNewLine & _
		"	}" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function Check(field)" & vbNewLine & _
		"    {" & vbNewLine & _
		"	field.checked = true;" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function Clear(field)" & vbNewLine & _
		"    {" & vbNewLine & _
		"	field.checked = false;" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function CheckAll()" & vbNewLine & _
		"    {" & vbNewLine & _
		"	var dm = document.delMembers;" & vbNewLine & _
		"	var len = dm.elements.length;" & vbNewLine & _
		"	for (var i = 0; i < len; i++) {" & vbNewLine & _
		"	    var field = dm.elements[i];" & vbNewLine & _
		"	    if (field.name == ""id"") {" & vbNewLine & _
		"		Check(field);" & vbNewLine & _
		"	    }" & vbNewLine & _
		"	}" & vbNewLine & _
		"	dm.toggleAll.checked = true;" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function ClearAll()" & vbNewLine & _
		"    {" & vbNewLine & _
		"	var dm = document.delMembers;" & vbNewLine & _
		"	var len = dm.elements.length;" & vbNewLine & _
		"	for (var i = 0; i < len; i++) {" & vbNewLine & _
		"	    var field = dm.elements[i];" & vbNewLine & _
		"	    if (field.name == ""id"") {" & vbNewLine & _
		"		Clear(field);" & vbNewLine & _
		"	    }" & vbNewLine & _
		"	}" & vbNewLine & _
		"	dm.toggleAll.checked = false;" & vbNewLine & _
		"    }" & vbNewLine & _
		"    function AllChecked()" & vbNewLine & _
		"    {" & vbNewLine & _
		"	dm = document.delMembers;" & vbNewLine & _
		"	len = dm.elements.length;" & vbNewLine & _
		"	for(var i = 0 ; i < len ; i++) {" & vbNewLine & _
		"	    if (dm.elements[i].name == ""id"" && !dm.elements[i].checked) {" & vbNewLine & _
		"		return false;" & vbNewLine & _
		"	    }" & vbNewLine & _
		"	}" & vbNewLine & _
		"	return true;" & vbNewLine & _
		"    }" & vbNewLine & _
		"    //-->" & vbNewLine & _
		"    </script>" & vbNewLine

selID = Request.QueryString("id")
strAction = Request.QueryString("action")
if strAction = "approve" then
	if selID = "-1" then
		Call EmailMembers("all")
		
		'## Forum_SQL - Approve all members
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING"
		strSql = strSql & " SET M_APPROVE = " & 1
		
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		
		Response.Write	"      <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Members Approved!</b></font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""5; URL=admin_accounts_pending.asp"">" & vbNewLine & _
				"      " & strParagraphFormat1 & "All Pending Members have been approved! Their registration e-mails have been sent to them.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""admin_accounts_pending.asp"">Back To Members Pending</font></a></p>" & vbNewLine
		WriteFooter
		Response.End
	else
		Call EmailMembers("selected")
		
		aryID = split(selID, ",")	
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Approve all members
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING"
			strSql = strSql & " SET M_APPROVE = " & 1
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)
			
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		next
		
		Response.Write	"      <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Members Approved!</b></font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""5; URL=admin_accounts_pending.asp"">" & vbNewLine & _
				"      " & strParagraphFormat1 & "Selected Pending Members have been approved! Their registration e-mails have been sent to them.</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""admin_accounts_pending.asp"">Back To Members Pending</font></a></p>" & vbNewLine
		WriteFooter
		Response.End
	end if
elseif strAction = "delete" then
	if selID = "-1" then
		'## Forum_SQL - Delete the Member
		strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " WHERE M_STATUS = " & 0
		strSql = strSql & " AND M_LEVEL = " & -1
		
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		
		Response.Write	"      <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Members Deleted!</b></font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=admin_accounts_pending.asp"">" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>All pending members have been deleted!</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""admin_accounts_pending.asp"">Back To Members Pending</font></a></p>" & vbNewLine
		WriteFooter
		Response.End
	
	else
		aryID = split(selID, ",")	
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Delete the Member
			strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)

			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		next
		Response.Write	"      <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """><b>Members Deleted!</b></font></p>" & vbNewLine & _
				"      <meta http-equiv=""Refresh"" content=""2; URL=admin_accounts_pending.asp"">" & vbNewLine & _
				"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Selected members have been deleted!</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<a href=""admin_accounts_pending.asp"">Back To Members Pending</font></a></p>" & vbNewLine
		WriteFooter
		Response.End
	end if
end if

mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)

'## Forum_SQL - Find all records with the search criteria in them
strSql = "SELECT M_NAME, M_EMAIL, MEMBER_ID, M_DATE, M_IP, M_KEY, M_APPROVE"
strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
strSql3 = " ORDER BY MEMBER_ID ASC;"

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		OffSet = cLng((mypage - 1) * strPageSize)
		strSql4 = " LIMIT " & OffSet & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "

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
			arrMemberData = rs.GetRows(intGetRows)
			iMemberCount = UBound(arrMemberData, 2)
		rs.close
		set rs = nothing
	else
		iMemberCount = ""
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
			arrMemberData = rs.GetRows(strPageSize)
			iMemberCount = UBound(arrMemberData, 2)
		else
			iMemberCount = ""
		end if
	rs.Close
	set rs = nothing
end if

Response.Write	"      <table border=""0"" align=center width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Members&nbsp;Pending...<br /><br /></font></td>" & vbNewLine
if maxpages > 1 then
	Response.Write	"          <td align=""right"">" & vbNewLine & _
			"            <table border=""0"" align=""right"">" & vbNewLine & _
			"              <tr>" & vbNewLine
	Call DropDownPaging(1)
	Response.Write	"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine
else
	Response.Write	"          <td align=""right"">&nbsp;</td>" & vbNewLine
end if
Response.Write	"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if iMemberCount <> "" then
	if strRestrictReg = "1" then scolspan = " colspan=""2"""
	Response.Write	"      <table border=""0"" cellSpacing=""0"" cellPadding=""0"" align=""center"">" & vbNewLine & _
  			"        <tr>" & vbNewLine & _
			"	   <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
			"	     <table align=""center"" width=""100%"" cellspacing=""1"" cellpadding=""4"" border=""0"">" & vbNewLine & _
	  		"	       <tr>" & vbNewLine & _
		    	"	         <td bgColor=""" & strHeadCellColor & """" & scolspan & "><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Administrator Options:</font></b></td>" & vbNewLine & _
	  		"	       </tr>" & vbNewLine & _
	  		"	       <tr>" & vbNewLine
	if strRestrictReg = "1" then
    		Response.Write	"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
				"                <li><a href=""javascript:appr_all()"">Approve All Pending Members</a></li>" & vbNewLine & _
				"                <li><a href=""javascript:appr_selected()"">Approve Selected Pending Members</a></li></font></td>" & vbNewLine
	end if
    	Response.Write	"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
			"                <li><a href=""javascript:del_all()"">Delete All Pending Members</a></li>" & vbNewLine & _
			"                <li><a href=""javascript:del_selected()"">Delete Selected Pending Members</a></li></font></td>" & vbNewLine & _
	  		"	       </tr>" & vbNewLine & _
			"	     </table>" & vbNewLine & _
			"	   </td>" & vbNewLine & _
  			"        </tr>" & vbNewLine & _
			"      </table><br />" & vbNewLine
end if

Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><b>NOTE:</b> The following table will show you a list of registered users that are waiting to be authenticated.</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine
if iMemberCount <> "" then
	Response.Write	"              <form name=""delMembers"" action=""admin_accounts_pending.asp"">" & vbNewLine & _
			"              <input type=""hidden"" name=""action"" value=""none"">" & vbNewLine
end if
Response.Write	"              <tr>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>User Name</font></b></td>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>E-mail Address</font></b></td>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Registered</font></b></td>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Days Since</font></b></td>" & vbNewLine & _
		"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Action</font></b></td>" & vbNewLine
if strRestrictReg = "1" then
		Response.Write "                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Approved?</font></b></td>" & vbNewLine
end if
Response.Write	"                <td bgColor=""" & strHeadCellColor & """ align=""center""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>"
if iMemberCount <> "" then
	Response.Write	"<input type=""checkbox"" name=""toggleAll"" value="""" onClick=""ToggleAll(this);"">"
else
	Response.Write	"&nbsp;"
end if
Response.Write	"</font></b></td>" & vbNewLine & _
		"              </tr>" & vbNewLine

if iMemberCount = "" then  '## No members found in DB
	Response.Write	"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strForumCellColor & """ colspan=""7""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><b>No Members Found</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine
else
	mM_NAME = 0
	mM_EMAIL = 1
	mMEMBER_ID = 2
	mM_DATE = 3
	mM_IP = 4
	mM_KEY = 5
	mM_APPROVE = 6

	rec = 1
	intI = 0

	for iMember = 0 to iMemberCount
		if (rec = strPageSize + 1) then exit for

		MP_MemberName = arrMemberData(mM_NAME, iMember)
		MP_MemberEMail = arrMemberData(mM_EMAIL, iMember)
		MP_MemberID = arrMemberData(mMEMBER_ID, iMember)
		MP_MemberDate = arrMemberData(mM_DATE, iMember)
		MP_MemberIP = arrMemberData(mM_IP, iMember)
		MP_MemberKey = arrMemberData(mM_KEY, iMember)
		MP_MemberApprove = arrMemberData(mM_APPROVE, iMember)

		if intI = 1 then 
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if
		
		if MP_MemberApprove = 1 then
			Approved = "Yes"
		else
			Approved = "No"
		end if

		days = DateDiff("d",  ChkDate(MP_MemberDate,"",false),  strForumTimeAdjust)
		if days >= 15 then
			days2 = "<b>" & days & "</b>"
		else
			days2 = days
		end if
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & chkString(MP_MemberName, "display") & "</a></font></td>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & MP_MemberEMail & "</font></td>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & ChkDate(MP_MemberDate,"",true) & "</font></td>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color="""
		if days >= 7 then Response.Write(strHiLiteFontColor) else Response.Write(strForumFontColor)
		Response.Write	""">" & days2 & "</font></td>" & vbNewLine & _
				"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText""><a href=""register.asp?actkey=" & MP_MemberKey & """>Activate Account</a></span></font></td>" & vbNewLine
		if strRestrictReg = "1" then
			Response.Write	"                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & Approved & "</font></td>" & vbNewLine
		end if
		Response.Write "                <td bgcolor=""" & CColor & """ align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><input type=""checkbox"" name=""id"" value=""" & MP_MemberID & """ onclick=""Toggle(this)""></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
		rec = rec + 1
		intI = intI + 1
		if intI = 2 then
			intI = 0
		end if
	next
	Response.Write	"              </form>"
end if
Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
if maxpages > 1 then
	Response.Write	"      <table border=""0"" align=""left"">" & vbNewLine & _
			"        <tr>" & vbNewLine
	Call DropDownPaging(2)
	Response.Write	"        </tr>" & vbNewLine & _
			"      </table><br />" & vbNewLine
else
	Response.Write	"      <br />" & vbNewLine
end if

WriteFooter
Response.End

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.write	"                <form name=""PageNum" & fnum & """ action=""admin_accounts_pending.asp"">" & vbNewLine
		Response.Write	"                <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine
		if fnum = 1 then
			Response.Write("                <b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		else
			Response.Write("                <b>There are " & maxpages & " Pages of Pending Members: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & vbNewLine)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then   
				Response.Write "                	<option value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			else
				Response.Write "                	<option selected value=""" & counter &  """>" & counter & "</option>" & vbNewLine
			end if
		next
		if fnum = 1 then
			Response.Write("                </select><b> of " & maxPages & "</b>" & vbNewLine)
		else
			Response.Write("                </select>" & vbNewLine)
		end if
		Response.Write("                </font></td>" & vbNewLine)
		Response.Write("                </form>" & vbNewLine)
	end if
end sub

sub EmailMembers(who)
	if who = "all" then
		'## Forum_SQL - Get all pending members
		strSql = "SELECT M_NAME, M_EMAIL, M_KEY, M_APPROVE"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
		strSql = strSql & " ORDER BY MEMBER_ID ASC"

		set rsApprove = Server.CreateObject("ADODB.Recordset")
		rsApprove.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

		if rsApprove.EOF then
			recApproveCount = ""
		else
			allApproveData = rsApprove.GetRows(adGetRowsRest)
			recApproveCount = UBound(allApproveData, 2)
		end if

		rsApprove.Close
		set rsApprove = Nothing

		if recApproveCount <> "" then
			mM_NAME = 0
			mM_EMAIL = 1
			mM_KEY = 2
			mM_APPROVE = 3

			for RowCount = 0 to recApproveCount
				MP_MemberName = allApproveData(mM_NAME,RowCount)
				MP_MemberEMail = allApproveData(mM_EMAIL,RowCount)
				MP_MemberKey = allApproveData(mM_KEY,RowCount)
				MP_MemberApprove = allApproveData(mM_APPROVE,RowCount)

				if MP_MemberApprove = "0" then
					'## E-mails Message to all pending members.
					strRecipientsName = MP_MemberName
					strRecipients = MP_MemberEMail
					strFrom = strSender
					strFromName = strForumTitle
					strsubject = strForumTitle & " Registration "
					strMessage = "Hello " & MP_MemberName & vbNewline & vbNewline
					strMessage = strMessage & "You received this message from " & strForumTitle & " because you have registered for a new account which allows you to post new messages and reply to existing ones on the forums at " & strForumURL & vbNewline & vbNewline
					if strAuthType="db" then
						strMessage = strMessage & "Please click on the link below to complete your registration." & vbNewline & vbNewLine
						strMessage = strMessage & strForumURL & "register.asp?actkey=" & MP_MemberKey & vbNewline & vbNewline
					end if
					strMessage = strMessage & "You can change your information at our website by selecting the ""Profile"" link." & vbNewline & vbNewline
					strMessage = strMessage & "Happy Posting!"
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
				end if
			next
		end if
	elseif who = "selected" then
		aryID = split(selID, ",")
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Get all pending members
			strSql = "SELECT M_NAME, M_EMAIL, M_KEY, M_APPROVE"
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)

			set rsApprove = my_Conn.Execute(strSql)

			if not(rsApprove.EOF) and not(rsApprove.BOF) and rsApprove("M_APPROVE") = "0" then
				'## E-mails Message to all pending members.
				strRecipientsName = rsApprove("M_NAME")
				strRecipients = rsApprove("M_EMAIL")
				strFrom = strSender
				strFromName = strForumTitle
				strsubject = strForumTitle & " Registration "
				strMessage = "Hello " & rsApprove("M_NAME") & vbNewline & vbNewline
				strMessage = strMessage & "You received this message from " & strForumTitle & " because you have registered for a new account which allows you to post new messages and reply to existing ones on the forums at " & strForumURL & vbNewline & vbNewline
				if strAuthType="db" then
					strMessage = strMessage & "Please click on the link below to complete your registration." & vbNewline & vbNewLine
					strMessage = strMessage & strForumURL & "register.asp?actkey=" & rsApprove("M_KEY") & vbNewline & vbNewline
				end if
				strMessage = strMessage & "You can change your information at our website by selecting the ""Profile"" link." & vbNewline & vbNewline
				strMessage = strMessage & "Happy Posting!"
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
				rsApprove.movenext
			end if
			rsApprove.Close
			set rsApprove = nothing
		next
	end if
end sub
%>
