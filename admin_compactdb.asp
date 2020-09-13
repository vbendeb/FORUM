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
Response.Write	"      <table border=""0"" width=""100%"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""33%"" align=""left"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All&nbsp;Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_forums.asp"">Forum&nbsp;Deletion/Archival</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Compact&nbsp;Database<br /></font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine

strForumDB = getForumDB()
strForumDB = replace(strForumDB,";","",1,1)
strDBPath = Left(strForumDB,InStrRev(strForumDB,"\"))
strTempFile = strDBPath & "Snitz_compacted.mdb"
DBFolderExists = CheckDBFolder(strDBPath)
if Application(strCookieURL & "down") then
	status = "Closed"
else
	status = "Open"
end if

Response.Write	"      <table border=""0"" width=""70%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Administrative Forum Archive Functions - Compact DB</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """ valign=""top""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strDefaultFontColor & """>" & vbNewLine
if request.querystring("action") = ""  then
	Response.Write	"                <p align=""center""><b>Depending on security settings at your Host, these operations may or may not be successful.  However, no harm should befall your data</b></p>" & vbNewLine & _
			"                Your original database will be copied to:<br /><br /><b>" & left(strForumDB,len(strForumDB)-4) & "_" & DateToStr(strForumTimeadjust) & ".bak" & "</b><br /><br />as a backup and then compacted to:<br /><br /><b>" & strTempFile & "</b><br />" & vbNewLine & _
			"                <br />" & vbNewLine & _
			"                If these steps are successful, the original DB will be replaced by the compacted DB.<br />" & vbNewLine & _
			"                <br />" & vbNewLine & _
			"                This may take some time depending on the size of your database.<br />" & vbNewLine & _
			"                <br />" & vbNewLine & _
			"                <font color=""" & strHiLiteFontColor & """><p align=""center"">You will have to <b>CLOSE</b> the forum while the database is being compacted.</font><br /><br />" & vbNewLine & _
			"                Current Status of Forum:<br /><b>" & status & "</b><br /><br />" & vbNewLine
	if Application(strCookieURL & "down") then
		Response.Write	"                Are you sure you want to compact the database?<br />" & vbNewLine & _
				"                <span class=""spnMessageText""><a href=""admin_compactdb.asp?action=Yes"">Yes</a></span>&nbsp;<span class=""spnMessageText""><a href=""admin_compactdb.asp?action=No"">No</a></span></p>" & vbNewLine
	else
		Response.Write	"                <span class=""spnMessageText""><a href=""down.asp?target=admin_compactdb.asp"">Click here</a></span> to close the forum before you start.</p>" & vbNewLine
	end if
elseif request.querystring("action") = "No" then
	Response.Write	"                <p align=""center"">You have chosen not to compact your database. You can compact your database at a later time.<br /><br />" & vbNewLine & _
			"                You will need to open your forums before you continue.<br />" & vbNewLine & _
			"                <span class=""spnMessageText""><a href=""down.asp?target=admin_forums.asp"">Click here</a></span> to open your forum.</p>" & vbNewLine
elseif request.querystring("action") = "Yes" then
	my_conn.close
	strTempConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTempFile
	
	if DBFolderExists = false then
		Response.Write	"                <br /><p align=""center"">Unable to create folder:<br />" & strDBPath & ".</p><br />" & vbNewLine
	else
		if BackupDB(strForumDB) then
			set jro = server.createobject("jro.JetEngine")
			jro.CompactDatabase strConnString, strTempConnString
			if err <> 0 then
				bError = True
				Response.Write	"                Error Compacting:<br />" & err.description & vbNewLine
			else
				Response.Write	"                <br /><p align=""center"">Database Compacted successfully.</p>" & vbNewLine
			end if
			if not bError then
				if not RenameFile( strTempFile, strForumDB) then
					Response.Write	"                Error Replacing:<br />" & err.description & vbNewLine
				else
					Response.Write	"                <br /><p align=""center"">Database renamed successfully.</p><br />" & vbNewLine
				end if
			end if
		else
			Response.Write	"                <br /><p align=""center"">Unable to back up database</p><br />" & vbNewLine
		end if
	end if
	set my_Conn = Server.CreateObject("ADODB.Connection")

	my_Conn.Open strConnString
	Response.Write	"                <p align=""center""><span class=""spnMessageText""><a href=""down.asp?target=admin_forums.asp"">Re-open Forum</a></span></p><br />" & vbNewLine
end if 
Response.Write	"                </font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"      <br />" & vbNewLine
'if Application(strCookieURL & "down") then
	'Response.Write	"      <p align=""center""><a href=""down.asp?target=admin_forums.asp""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strDefaultFontColor & """ >Open Forum</font></a></p>" & vbNewLine
	Response.Write	"      <p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strDefaultFontColor & """><a href=""admin_forums.asp"">Back to Forums Administration</a></font></p>" & vbNewLine
'end if
Response.Write	"      <br />" & vbNewLine & _
		"      <br />" & vbNewLine
WriteFooter
Response.End

Function RenameFile(sFrom, sTo)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		RenameFile = False
		Exit Function
	end if	
	fso.DeleteFile sTo,true
	fso.MoveFile sFrom, sTo
	set fso = nothing
	RenameFile = True
end Function

Function BackupDB(sFrom)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		BackupDB = False
		Exit Function
	end if	

	fso.CopyFile sFrom, fso.GetParentFolderName(sFrom) & "\" & fso.GetBaseName(sFrom) & "_" & DateTostr(strForumTimeAdjust) & ".bak", true
	set fso = nothing
	BackupDB = True
end Function

Function GetForumDB()
	dim tmpFileName
 	tmpFileName = split(strConnstring,"Source=",2,1)
	GetForumDB = tmpFileName(1)
end Function

Function CheckDBFolder(strPath)
	Dim fso, blnExists
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		CheckDBFolder = False
		Exit Function
	end if
	
	blnExists = fso.FolderExists(strPath)
	if blnExists = false then
		fso.CreateFolder(strPath)
		CheckDBFolder = True
	else
		CheckDBFolder = True
	end if
End Function
%>
