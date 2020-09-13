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
<!--#include file="config.asp"-->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
%>
<!--#include file="inc_header.asp"-->
<%
if MemberID <> intAdminMemberID then
	Err_Msg = "<li>Only the Forum Admin can access this page</li>"

	Response.Write	"<table align=""center"" width=""50%"" height=""50%"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td>" & vbNewLine & _
			"    <p align=""center""><font face=""Verdana, Arial, Helvetica"" size=""3"" color=""#FF0000"">There has been a problem!</font></p>" & vbNewLine & _
			"      <table align=""center"" border=""0"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td><font face=""Verdana, Arial, Helvetica"" size=""2"" color=""#FF0000""><ul>" & Err_Msg & "</ul></font></td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"    <p align=""center"" valign=""middle""><font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""JavaScript:history.go(-1)"">Go Back To Admin Section</a></font></p>" & vbNewLine & _
			"    </td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
	Response.End
end if

Response.Write	"<div align=""center""><center><p><font face=""Verdana, Arial, Helvetica"" size=""4"">" & _
		"Snitz Forum Modifications</font></p></center></div>" & vbNewLine

Dim strTableName
Dim fieldArray (100)
Dim idFieldName
Dim tableExists
Dim fieldExists
Dim ErrorCount

tableExists   = -2147217900
tableNotExist = -2147217865 
fieldExists   = -2147217887
ErrorCount = 0
	
on error resume next
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if err.number <> 0 then
	response.write "error " & err.number & "|" & err.description
	response.redirect "admin_mod_dbsetup2.asp"
	err.clear
	response.end
end if

set objFile = fso.Getfile(server.mappath(Request.ServerVariables("PATH_INFO")))
set objFolder = objFile.ParentFolder
set objFolderContents = objFolder.Files

if Request.Form("dbMod") = "" then

	Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <b>Database Setup....</b><br />"
	If strDBType = "" then 
		Response.Write	"<font face=""Verdana, Arial, Helvetica"" color=""#FF0000"" size=""2"">Your strDBType is not set, please edit your config.asp<br />" & _
				"to reflect your database type<br /></font>" & _
				"<br /><a href=""default.asp"">Go Back to Forum</a></font>"
		Response.End
	end if
	Response.Write	"    <form action=""" & Request.ServerVariables("PATH_INFO") & """ method=""post"" name=""form1"">" & vbNewLine
	if strDBType = "sqlserver" then 
		Response.Write	"    <font face=""Verdana, Arial, Helvetica"" size=""1"">" & _
				"You are using SQL Server, please select the correct version<br />" & vbNewLine & _
				"    <input type=""radio"" name=""sqltype"" value=""7"" checked> SQL 7.x&nbsp;&nbsp;&nbsp;&nbsp;" & vbNewLine & _
				"    <input type=""radio"" name=""sqltype"" value=""6""> SQL 6.x<br /></font>" & vbNewLine
	end if

	on error resume next
	Response.Write	"    <font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine & _
			"    <p>Select the Mod from the list below, and press Update!<br />" & vbNewLine & _
			"    A script will execute to perform the database upgrade.</p></font>" & vbNewLine & _
			"    <select name=""dbMod"" size=""1"">" & vbNewLine
	for each objFileItem in objFolderContents
		intFile = instr(objFileItem.Name, "dbs_")
    	if intFile <> 0 then
        	whichfile = server.mappath(objFileItem.Name)
	    	Set fs = CreateObject("Scripting.FileSystemObject")
        	Set thisfile = fs.OpenTextFile(whichfile, 1, False)
			ModName = thisfile.readline
			Response.Write	"    	<option value=""" & whichfile & """>" & ModName & "</option>"
			thisfile.close
			if err.number <> 0 then 
				Response.Write err.description
				Response.end
			end if
			set fs = nothing
  		end if
	Next
	Response.Write	"    </select>" & vbNewLine & _
			"    <input type=""submit"" name=""submit1"" value=""Update!""><br />" & vbNewLine & _
			"    <input type=""checkbox"" name=""delFile"" value=""1"">Delete the dbs file when finished?</form>" & vbNewLine & _
			"    </font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center"">" & vbNewLine & _
			"    <font face=""Verdana, Arial, Helvetica"" size=""2""><a href=""default.asp"" target=""_top"">Click here to go to the forum.</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine
else
	Response.Write	"<table border=""0"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" align=""center"">" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td bgColor=""#9FAFDF"" align=""center"">" & vbNewLine & _
			"    <p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine
	sqlVer = Request.Form("sqltype")
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set thisfile = fs.OpenTextFile(Request.Form("dbMod"), 1, False)
	ModName = thisfile.readline
	response.write ("    <font face=""Verdana, Arial, Helvetica"" size=""3"">")
	response.write ("    <h4>" & ModName & "</h4></font>")

	'## Load Sections for processing
	do while not thisfile.AtEndOfStream
		sectionName = thisfile.readline
		Select case uCase(sectionName)
			case "[CREATE]" 
				strTableName   =   uCase(thisfile.readline)
				idFieldName = uCase(thisfile.readline)
				tempField = thisfile.readline
				rec = 0
				do while uCase(tempField) <>  "[END]"
					fieldArray(rec) = tempField
					rec = rec+1
					tempField = thisfile.readline
				loop
				CreateTables(rec)
			case "[ALTER]" 
				strTableName   =   uCase(thisfile.readline)
				tempField = thisfile.readline
				rec = 0
				do while uCase(tempField) <>  "[END]"
					fieldArray(rec) = tempField
					rec = rec+1
					tempField = thisfile.readline
				loop
				AlterTables(rec)
			case "[DELETE]" 
				strTableName   =   uCase(thisfile.readline)
				tempField = thisfile.readline
				rec = 0
				do while uCase(tempField) <>  "[END]"
					fieldArray(rec) = tempField
					rec = rec+1
					tempField = thisfile.readline
				loop
				DeleteValues(rec)
			case "[INSERT]" 
				strTableName   =   uCase(thisfile.readline)
				tempField = thisfile.readline
				rec = 0
				do while uCase(tempField) <>  "[END]"
					fieldArray(rec) = tempField
					rec = rec+1
					tempField = thisfile.readline
				loop
				InsertValues(rec)
			case "[UPDATE]" 
				strTableName   =   uCase(thisfile.readline)
				tempField = thisfile.readline
				rec = 0
				do while uCase(tempField) <>  "[END]"
					fieldArray(rec) = tempField
					rec = rec+1
					tempField = thisfile.readline
				loop
				UpdateValues(rec)
			case "[DROP]" 
				strTableName   =   thisfile.readline
				tempField = thisfile.readline
				DropTable()
		end select
	loop
	Response.Write	""
	if request("delFile") = "1" then
			thisfile.close
			on error resume next
			fs.DeleteFile(Request.Form("dbMod"))
			if err.number = 0 then
				Response.write "    <font face=""Verdana, Arial, Helvetica"" size=""2""><b>The dbs file was succesfully deleted.</b></font><br />" & vbNewLine
			else
				Response.write "    <font face=""Verdana, Arial, Helvetica"" size=""2""><b>Unable to remove dbs file<br /><font color=""#FF0000"">" & err.description & "</font></font>" & vbNewLine
			end if
	end if
	if ErrorCount > 0 then
		Response.write	"    <br />If there were errors please post a question in the MOD Implementation Forum at<br />" & vbNewLine & _
				"    <a href=""http://forum.snitz.com/forum/forum.asp?FORUM_ID=94"">Snitz Forums</a>" & vbNewLine
	else
		Response.Write	"    <br /><font face=""Verdana, Arial, Helvetica"" size=""2""><p><b>Database setup finished</b></p>" & vbNewLine
	end if
	Response.Write	"    </font>" & vbNewLine & _
			"    <form action=""" & Request.ServerVariables("PATH_INFO") & """ method=""post"" name=""form2"">" & vbNewLine & _
			"    <input type=""hidden"" name=""dbMod"" value="""">" & vbNewLine & _
			"    <input type=""submit"" name=""submit2"" value=""Finished""></form>" & vbNewLine & _
			"    </font></p></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"  <tr>" & vbNewLine & _
			"    <td align=""center""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & vbNewLine & _
			"    <a href=""default.asp"" target=""_top"">Click here to go to the forum.</a></font></td>" & vbNewLine & _
			"  </tr>" & vbNewLine & _
			"</table>" & vbNewLine & _
			"</form>" & vbNewLine
end if 

set fs = nothing
set fso = nothing
WriteFooter
Response.End

Sub CreateTables( numfields )
	response.write "    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	response.write "    <b>Creating table(s)...</b><br />" & vbNewLine
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		TablePrefix = strMemberTablePrefix
	else
		TablePrefix = strTablePrefix
	end if

	strSql = "CREATE TABLE " & TablePrefix & strTableName & "( "
	if idFieldName <> "" then
		select case strDBType
			case "access"
				if Instr(strConnString,"(*.mdb)") then
					strSql = strSql & idFieldName &" COUNTER CONSTRAINT PrimaryKey PRIMARY KEY "
				else
					strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
				end if
			case "sqlserver"
				strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
			case "mysql"
				strSql = strSql & idFieldName &" INT (11) DEFAULT '' NOT NULL auto_increment "
		end select
	end if
	for y = 0 to numfields -1
	on error resume next
		tmpArray = split(fieldArray(y),"#")
		fName = uCase(tmpArray(0))
		fType = lCase(tmpArray(1))
		fNull = uCase(tmpArray(2))
		fDefault = tmpArray(3)
		if idFieldName <> "" or y <> 0 then
			strSql = strSql & ", "
		end if
		select case strDBType
			case "access"
				fType = replace(fType,"varchar (","text (")
			case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
			case "mysql"
				fType = replace(fType,"memo","text")
				fType = replace(fType,"#int","#int (11)")
				fType = replace(fType,"#smallint","#smallint (6)")
		end select
		if fNull <> "NULL" then fNull = "NOT NULL"
		strSql = strSql & fName & " " & fType & " " & fNull & " " 
		if fdefault <> "" then
			select case strDBType
				case "access"
					if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
				case else
					strSql = strSql & "DEFAULT " & fDefault
			end select
		end if
	next
	if strDBType = "mysql" then
		if idFieldName <> "" then
			strSql = strSql & ",KEY " & TablePrefix & strTableName & "_" & idFieldName & "(" & idFieldName & "))"
		else
			strSql = strSql & ")"
		end if
	else
		strSql = strSql & ")"
	end if
	response.write "    " & strSql & "<br />" & vbNewLine
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	if err.number <> 0 and err.number <> 13 and err.number <> tableExists then
		response.Write "    " & strSql & "<br />" & vbNewLine
		response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableExists then 
			Response.Write("    <font color=""#FF0000""><b>Table already exists</b></font><br />" & vbNewLine)
		else
			Response.Write("    <b>Table created succesfully</b><br />" & vbNewLine)
		end if
	end if
	
	response.write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub

Sub AlterTables(numfields)
	Response.write "    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	for y = 0 to numfields -1
		on error resume next
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			TablePrefix = strMemberTablePrefix
		else
			TablePrefix = strTablePrefix
		end if
		strSql = "ALTER TABLE " & TablePrefix & strTableName 
		tmpArray = split(fieldArray(y),"#")
		fAction = uCase(tmpArray(0))
		fName = uCase(tmpArray(1))
		fType = lCase(tmpArray(2))
		fNull = uCase(tmpArray(3))
		fDefault = tmpArray(4)
		select case fAction
			case "ADD"
				strSQL = strSQL & " ADD "
				if strDBType = "access" then strSql = strSql & "COLUMN "
			case "DROP"
				strSQL = strSQL & " DROP COLUMN "
			case "ALTER"
				strSQL = strSQL & " ALTER COLUMN "
			case else
		end select
		if fAction = "ADD" or fAction = "ALTER" then
			select case strDBType
				case "access"
					fType = replace(fType,"varchar (","text (")
				case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
				case "mysql"
					fType = replace(fType,"memo","text")
					fType = replace(fType,"#int","#int (11)")
					fType = replace(fType,"#smallint","#smallint (6)")
			end select
			if fNull <> "NULL" then fNull = "NOT NULL"
			strSql = strSQL & fName & " " & fType & " " & fNULL & " "
			if fDefault <> "" then
				select case strDBType
					case "access"
						if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
					case else
						strSql = strSql & "DEFAULT " & fDefault
				end select
			end if
			Response.Write	"    <b>Adding Column " & fName & "...</b><br />" & vbNewLine
		else
			strSql = strSQL & fName
			Response.Write "    <b>Dropping Column...</b><br />" & vbNewLine
		end if
		response.write "    " & strSql & "<br />" & vbNewLine
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 and err.number <> 13 and err.number <> fieldExists then
			response.write "    " & strSQL & "<br />" & vbNewLine
			response.write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
			ErrorCount = ErrorCount + 1
			resultString = ""
		else
			if fAction = "DROP" then
				Response.Write("    <b>Column " & LCase(fAction) & "ped successfully</b><br />" & vbNewLine)
				resultString = "    <b>Table(s) updated</b><br />" & vbNewLine
			else
				if err.number = fieldExists then 
					Response.Write("    <b><font color=""#FF0000"">Column already exists</font></b><br />" & vbNewLine)
					resultString = ""
				else
					Response.Write("    <b>Column " & LCase(fAction) & "ed successfully</b><br />" & vbNewLine)
				end if
			end if
			if fDefault <> "" and err.number <> fieldExists then
				strSQL = "UPDATE " & TablePrefix & strTableName & " SET " & fName & "=" & fDefault
				response.write "    " & strSql & "<br />" & vbNewLine
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				response.write "    <b>Populating Current Records with new Default value</b><br />" & vbNewLine
				resultString = "    <b>Table(s) updated</b><br />" & vbNewLine
			end if
		end if
		
		if fieldArray(y) = "" then y = numfields
	next
	Response.Write(resultString)
	Response.Write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub

Sub InsertValues(numfields)
	Response.Write "    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	on error resume next
	Response.Write ("    <b>Adding new records..</b><br />" & vbNewLine)
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "INSERT INTO " & strMemberTablePrefix & strTableName & " "
		else
			strSql = "INSERT INTO " & strTablePrefix & strTableName & " "
		end if
		tmpArray = split(fieldArray(y),"#")
		fNames = tmpArray(0)
		fValues = tmpArray(1)
		strSql = strSql & tmpArray(0) & " VALUES " & tmpArray(1)
		Response.Write	"    " & strSql & "<br />" & vbNewLine
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	next

	if err.number <> 0 and err.number <> 13 then
		Response.Write "    " & strSql & "<br />" & vbNewLine
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		Response.Write("    <br /><b>Value(s) updated succesfully</b>" & vbNewLine)
	end if
	Response.Write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub 

Sub UpdateValues(numfields)
	on error resume next
	Response.write	"    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	response.write("    <b>Updating Forum Values..</b><br />" & vbNewLine)
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "UPDATE " & strMemberTablePrefix & strTableName & " SET"
		else
			strSql = "UPDATE " & strTablePrefix & strTableName & " SET"
		end if
		tmpArray = split(fieldArray(y),"#")
		fName = tmpArray(0)
		fValue = tmpArray(1)
		fWhere = tmpArray(2)
		strSql = strSql & " " & fName & " = " & fvalue
		if fWhere <> "" then
			strSql = strSql & " WHERE " & fWhere
		end if
		Response.Write "    " & strSql & "<br />" & vbNewLine
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	next

	if err.number <> 0 then
		Response.Write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
		Response.Write "    " & strSql & "<br />" & vbNewLine
	else
		Response.Write("    <br /><b>Value(s) updated succesfully</b>" & vbNewLine)
	end if
	Response.Write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub 

Sub DeleteValues(numfields)
	on error resume next
	response.write "    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	response.write("    <b>Updating Forum Values..</b><br />" & vbNewLine)
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DELETE FROM " & strMemberTablePrefix & strTableName & " WHERE "
	else
		strSql = "DELETE FROM " & strTablePrefix & strTableName & " WHERE "
	end if
	tmpArray = fieldArray(0)
	strSql = strSql & tmpArray
	response.write "    " & strSql & "<br />" & vbNewLine
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 then
		response.write "    " & strSql & "<br />" & vbNewLine
		response.write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		response.write("    <br /><b>Value(s) updated succesfully</b>" & vbNewLine)
	end if
	response.write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub 

Sub DropTable()
	on error resume next
	response.write "    <br /><font face=""Verdana, Arial, Helvetica"" size=""1"">" & vbNewLine
	response.write("    <b>Dropping Table..</b><br />" & vbNewLine)
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DROP TABLE " & strMemberTablePrefix & strTableName
	else
		strSql = "DROP TABLE " & strTablePrefix & strTableName
	end if
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 and err.number <> 13 and err.number <> tableNotExist then
		response.write "    " & strSql & "<br />" & vbNewLine
		response.write("    <font color=""#FF0000"">" & err.number & " | " & err.description & "</font><br />" & vbNewLine)
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableNotExist then
			response.write("    <br /><b>Table does not exist</b>" & vbNewLine)
		else
			response.write("    <br /><b>Table dropped succesfully</b>" & vbNewLine)
		end if
	end if
	response.write("    <hr size=""1"" width=""260"" align=""center"" color=""blue""></font>" & vbNewLine)
end Sub

on error goto 0
%>