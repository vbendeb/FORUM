<%Option Explicit%>
<%
Response.Write	"<html>" & vbNewLine & _
		"<head>" & vbNewLine & _
		"<title>Choose Your Birthdate</title>" & vbNewLine & _
		"<script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"// SET FORM FIELD VALUE TO THE DATE SELECTED" & vbNewLine & _
		vbNewLine & _
		"function ignore() {" & vbNewLine & _
		"	return true;" & vbNewLine & _
		"}" & vbNewLine & _
		"function returnDate(inDay) {" & vbNewLine & _
		"	window.onerror = ignore" & vbNewLine & _
		"	window.opener.document.forms['" & Request("FormName") & "']['" & Request("FieldName") & "'].value = inDay;" & vbNewLine & _
		"	self.close();" & vbNewLine & _
		"}" & vbNewLine & _
		"</script>" & vbNewLine

Rem -Get info from Application Variables
dim strCookieURL, strTimeAdjust, strForumTimeAdjust
strCookieURL = Left(Request.ServerVariables("Path_Info"), InstrRev(Request.ServerVariables("Path_Info"), "/"))
strTimeAdjust = Application(strCookieURL & "STRTIMEADJUST")
strForumTimeAdjust = DateAdd("h", strTimeAdjust , Date())
Rem -Color and Font vars
dim strDefaultFontFace,strDefaultFontSize,strHeaderFontSize,strFooterFontSize
dim strPageBGColor,strDefaultFontColor,strHeadCellColor,strHeadFontColor
dim strCategoryCellColor,strCategoryFontColor,strForumCellColor,strAltForumCellColor
dim strForumFontColor,strForumLinkColor,strForumLinkTextDecoration,strForumVisitedLinkColor
dim strForumVisitedTextDecoration,strForumActiveLinkColor,strForumActiveTextDecoration
dim strForumHoverFontColor,strForumHoverTextDecoration,strTableBorderColor,strHiLiteFontColor
dim strPageBGImageURL
strDefaultFontFace = Application(strCookieURL & "STRDEFAULTFONTFACE")
strDefaultFontSize = Application(strCookieURL & "STRDEFAULTFONTSIZE")
strHeaderFontSize = Application(strCookieURL & "STRHEADERFONTSIZE")
strFooterFontSize = Application(strCookieURL & "STRFOOTERFONTSIZE")
strPageBGColor = Application(strCookieURL & "STRPAGEBGCOLOR")
strDefaultFontColor = Application(strCookieURL & "STRDEFAULTFONTCOLOR")
strHeadCellColor = Application(strCookieURL & "STRHEADCELLCOLOR")
strHeadFontColor = Application(strCookieURL & "STRHEADFONTCOLOR")
strCategoryCellColor = Application(strCookieURL & "STRCATEGORYCELLCOLOR")
strCategoryFontColor = Application(strCookieURL & "STRCATEGORYFONTCOLOR")
strForumCellColor = Application(strCookieURL & "STRFORUMCELLCOLOR")
strAltForumCellColor = Application(strCookieURL & "STRALTFORUMCELLCOLOR")
strForumFontColor = Application(strCookieURL & "STRFORUMFONTCOLOR")
strForumLinkColor = Application(strCookieURL & "STRFORUMLINKCOLOR")
strForumLinkTextDecoration = Application(strCookieURL & "STRFORUMLINKTEXTDECORATION")
strForumVisitedLinkColor = Application(strCookieURL & "STRFORUMVISITEDLINKCOLOR")
strForumVisitedTextDecoration = Application(strCookieURL & "STRFORUMVISITEDTEXTDECORATION")
strForumActiveLinkColor = Application(strCookieURL & "STRFORUMACTIVELINKCOLOR")
strForumActiveTextDecoration = Application(strCookieURL & "STRFORUMACTIVETEXTDECORATION")
strForumHoverFontColor = Application(strCookieURL & "STRFORUMHOVERFONTCOLOR")
strForumHoverTextDecoration = Application(strCookieURL & "STRFORUMHOVERTEXTDECORATION")
strTableBorderColor = Application(strCookieURL & "STRTABLEBORDERCOLOR")
strHiLiteFontColor = Application(strCookieURL & "STRHILITEFONTCOLOR")

Response.Write	"<style>" & vbNewLine & _
		"<!--" & vbNewLine & _
		".spnMessageText a:link    {color:" & strForumLinkColor & ";text-decoration:" & strForumLinkTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:visited {color:" & strForumVisitedLinkColor & ";text-decoration:" & strForumVisitedTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:hover   {color:" & strForumHoverFontColor & ";text-decoration:" & strForumHoverTextDecoration & "}" & vbNewLine & _
		".spnMessageText a:active  {color:" & strForumActiveLinkColor & ";text-decoration:" & strForumActiveTextDecoration & "}" & vbNewLine & _
		"//-->" & vbNewLine & _
		"</style>" & vbNewLine & _
		"</head>" & vbNewLine & _
		"<body background=""" & strPageBGImageURL & """ bgColor=""" & strPageBGColor & """ text=""" & strDefaultFontColor & """ topmargin=""3"" marginheight=""3"" marginwidth=""8"">" & vbNewLine

Rem -You can pass in a date...   path/filename.asp?date=5/6/2003  It defaults to Todays date
Rem -To turn on a no select date after today pass in "History=on" in the url
Rem -Changed by Rakesh Jain(GauravBhabu)
function GetDaysInMonth(ByVal iMonth, ByVal iYear)
	dim arrDaysInMonth
	arrDaysInMonth = Array(31,28,31,30,31,30,31,31,30,31,30,31)
	if isLeapYear(iYear) then arrDaysInMonth(1) = 29
	GetDaysInMonth = arrDaysInMonth(iMonth -1)
end Function
Rem -This Procedure checks for leap year
Rem -Added by Rakesh Jain(GauravBhabu)
function IsLeapYear(ByVal intYear) 'As Integer) As Boolean
	IsLeapYear = False
	if (intYear Mod 100 = 0) then
		if (intYear Mod 400 = 0) then IsLeapYear = True
	elseif (intYear Mod 4 = 0) then
		IsLeapYear = True
	end if
end function
function GetWeekdayMonthStartsOn(ByVal dAnyDayInTheMonth)
	dim dTemp
	Rem -Deduct (Day Of Month - 1) from date to Get the date on first day of Month
	dTemp = DateAdd("d", -(Day(dAnyDayInTheMonth) - 1), dAnyDayInTheMonth)
	GetWeekdayMonthStartsOn = WeekDay(dTemp)
end function
Rem -Changed by Rakesh Jain(GauravBhabu)
function PreviousMonth(ByVal dDate)
	dim dtePrevMonth
	dtePrevMonth = DateAdd("m", -1, dDate)
	PreviousMonth = dtePrevMonth
end function
Rem -Changed by Rakesh Jain(GauravBhabu)
function NextMonth(ByVal dDate)
	dim dteNextMonth
	dteNextMonth = DateAdd("m", 1, dDate)
	if Month(dteNextMonth) > Month(strForumTimeAdjust) and Year(dteNextMonth) = Year(strForumTimeAdjust) then 
		dteNextMonth = strForumTimeAdjust
	end if
	NextMonth = dteNextMonth
end function
Rem -This procedure writes the days of month for the calendar
Rem -Added by Rakesh Jain(GauravBhabu)
sub WriteDayOfMonth(ByVal strDate, ByVal strClass, ByVal intDay, ByVal blnOnClick)
	Dim strDayLink, strOnClick, strCellColor, strBoxTitle
	strBoxTitle = ""
	if blnOnClick then 
		strBoxTitle = FormatdateTime(dCell,vbLongdate)
		strOnClick =  " onclick=""" & strReturnFunc & """"
		if strClass = "" then 
			strCellColor = strForumFontColor
			strDayLink =	"<a href=""javascript:" & strReturnFunc & """><font color=""" & strForumCellColor & """><b>" & intDay & "</b></font></a>"
		else
			strCellColor = strForumCellColor
			strDayLink =	"<span class=""" & strClass & """><a href=""javascript:" & strReturnFunc & """><b>" & intDay & "</b></a></span>"
		end if
	else
		strCellColor = strAltForumCellColor
		strDayLink = intDay
	end if
	Response.Write	"          <td title=""" & strBoxTitle & """ bgcolor=""" & strCellColor & """" & strOnClick & "><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & strDayLink & "<br /><br /></font></td>" & vbNewLine
end sub
function SetMonthSelection(i_intMonth)
	if i_intMonth = Month(dDate) then
		SetMonthSelection = (" selected")
	end if
end function
sub SetYearSelection(i_intYear)
	if i_intYear = Year(dDate) then
		SetYearSelection = (" selected")
	end if
end sub
Rem -Append zeros to the left of single digit Months and days
function doublenum(fNum)
	if fNum > 9 then
		doublenum = fNum
	else
		doublenum = "0" & fNum
	end if
end function
function IsValidDate(strDOBDate)
	dim intYear, intMonth, intDay
	IsValidDate = false
	if IsNumeric(strDOBDate) then
		if len(strDOBDate) = 8 then
			intYear = cLng(Left(strDOBDate,4))
			intMonth = clng(Mid(strDOBDate,5,2))
			intDay = cLng(Mid(strDOBDate,7,2))
			if IsValidYear(intYear) then
				if IsValidMonth(intMonth) then
					if IsValidDay(intYear,intMonth,intDay) then IsValidDate = true
				end if
			end if
		end if
	end if
end function
function IsValidYear(ByVal intYear)
	IsValidYear = false
	if (intYear > 1900) and (intYear <= Year(Date)) then IsValidYear = true
end function
function IsValidMonth(ByVal intMonth)
	IsValidMonth = false
	if intMonth > 0 and intMonth < 13 then IsValidMonth = true
end function
function IsValidDay(ByVal intYear,ByVal intMonth,ByVal intDay)
	dim arrDaysInMonth
	arrDaysInMonth = Array(31,28,31,30,31,30,31,31,30,31,30,31)
	IsValidDay = false
	if IsLeapYear(intYear) then arrDaysInMonth(1) = 29
	if (intDay) <= arrDaysInMonth(intMonth-1) then IsValidDay = true
end function
Rem -End Function Declaration
dim dDate : Rem -Date we're displaying calendar for
dim iDIM : Rem -Days In Month
dim iDOW : Rem -Day Of Week that month starts on
dim iCurrent : Rem -Variable we use to hold current day of month as we write table
dim iPosition : Rem -Variable we use to hold current position in table
dim strDOBDate : Rem -Holds the date of Birth if there is one - YYYYMMDD
dim strReturnFuncEmpty
Rem -Get selected date.  There are two ways to do this.
Rem -First check if we were passed a full date in RQS("date").
Rem -If so use it, if not look for seperate variables, putting them togeter into a date.
Rem -Lastly check if the date is valid...if not use today
if IsDate(Request.QueryString("date")) then
	Rem -This is date when navigating the calendar
	Rem -This should be a date as per locale Format
	dDate = cDate(Request.QueryString("date"))
elseif IsValidDate(Request.QueryString("date")) then
	Rem -This is when user edits Date of Birth 
	Rem -This should be in YYYYMMDD Format
	strDOBDate = Request.QueryString("date")
	dDate = cDate(Mid(strDOBDate,7,2) & "-" & MonthName(Mid(strDOBDate, 5,2)) & "-" & Mid(strDOBDate, 1,4))
else   '****************** Put as one ***********
	Rem -Assign a Default date to dDate variable
	dDate = DateAdd("yyyy", -13, strForumTimeAdjust)
	dDate = DateValue(dDate)
	if Request("day") <> "" and Request("month") <> "" and Request("year") <> "" then
		Rem -This will be the date when User clicks on Go Button
		if IsDate(Request("day") & "-" & MonthName(Request("month")) & "-" & Request("year")) Then
			dDate = cDate(Request("day") & "-" & MonthName(Request("month")) & "-" & Request("year"))
		end if
	end if
end if	

Response.Write	"<form action=""" & Request.ServerVariables("PATH_INFO") & "?FormName=" & Request("FormName") & "&FieldName=" & Request("FieldName") & "&History=" & Request("History") & """ method=""post"" id=""form1"" name=""form1"">" & vbNewLine & _
		"<input type=""hidden"" name=""day"" value=""" & Day(dDate) & """>" & vbNewLine & _
		"<table width=""275"" border=""0"" cellspacing=""1"" cellpadding=""2"" align=""center"">" & vbNewLine & _
		"  <tr>" & vbNewLine & _
		"    <td align=""center"">" & vbNewLine
Rem -Month Select Box
dim intLastMonth
Rem -Restrict the available Dates to Today
if Year(dDate) = Year(strForumTimeAdjust) then
	intLastMonth = Month(strForumTimeAdjust)
else
	intLastMonth = 12
end if
if Month(dDate) >= intLastMonth and Year(dDate) >= Year(strForumTimeAdjust) then
	if Day(dDate) > Day(strForumTimeAdjust) then
		dDate = DateSerial(Year(dDate),intLastMonth, 1) 
	else
		dDate = DateSerial(Year(dDate),intLastMonth, Day(dDate))
	end if
end if
Rem -Days in Month
iDIM = GetDaysInMonth(Month(dDate), Year(dDate))
Rem -Day of Week on First of Month
iDOW = GetWeekdayMonthStartsOn(dDate)
dim iMonth 'Counter to fill the Month Select Box
Response.Write	"    <select name=""month"">" & vbNewLine
for iMonth = 1 to intLastMonth
	Response.Write	"    	<option value=""" & iMonth & """" & SetMonthSelection(iMonth) & ">" & MonthName(iMonth) & "</option>" & vbNewLine
next
Response.Write	"    </select>" & vbNewLine
Rem -Year Select Box
dim int_YearCntr
Response.Write	"    <select name=""year"">" & vbNewLine
for int_YearCntr = 1901 to Year(strForumTimeAdjust) step 1
	Response.Write	"    	<option value=""" & int_YearCntr & """"
	if int_YearCntr = Year(dDate) then
		Response.Write(" selected")
	end if
	Response.Write	">" & int_YearCntr & "</option>" & vbNewLine
next
Response.Write	"    </select>" & vbNewLine & _
		"    <input type=""submit"" VALUE=""Go"" id=""submit1"" NAME=""submit1"">" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine
Rem -Calendar Navigation
Response.Write	"<table width=""275"" border=""0"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"  <tr>" & vbNewLine & _
		"    <td bgcolor=""" & strHeadCellColor & """ align=""center"">" & vbNewLine & _
		"      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td width=""15%"" align=""right""><a href=""pop_datepicker.asp?date=" & PreviousMonth(dDate) & "&FormName=" & Request("FormName") & "&FieldName=" & Request("FieldName") & "&History=" & Request("History") & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&lt;&lt;</font></a></td>" & vbNewLine & _
		"          <td align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """><b>" & MonthName(Month(dDate)) & "  " & Year(dDate) & "</b></font></td>" & vbNewLine & _
		"          <td width=""15%"" align=""left"">"
if NextMonth(dDate) <> strForumTimeAdjust then Response.Write("<a href=""pop_datepicker.asp?date=" & NextMonth(dDate) & "&FormName=" & Request("FormName") & "&FieldName=" & Request("FieldName") & "&History=" & Request("History") & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>&gt;&gt;</font></a>") else Response.Write("&nbsp;")
Response.Write	"</td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine
Rem -Weekday Names
Response.Write	"  <tr>" & vbNewLine & _
		"    <td bgcolor=""" & strTableBorderColor & """>" & vbNewline & _
		"      <table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""2"" align=""center"">" & vbNewline & _
		"        <tr>" & vbNewLine
dim iWeekDayName
for iWeekDayName = 1 to 7
	Response.Write	"          <td width=""14%"" align=""center"" bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strCategoryFontColor & """>" & WeekDayName(iWeekDayName, True) & "</font></td>" & vbNewLine
next
Response.Write	"        </tr>" & vbNewLine
strReturnFuncEmpty = "returnDate(' '); "
Rem -Write spacer cells at beginning of first row if month doesn't start on a Sunday.
if iDOW <> 1 then
	iPosition = iDOW
	Response.Write	"        <tr>" & vbNewLine & _
			"          <td colspan=""" & iPosition - 1 & """ bgcolor=""#bbbbbb""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>&nbsp;<br /><br /></font></td>" & vbNewLine
end if
Rem -Write days of month in proper day slots
Dim strReturnFunc, dCell, strClass, blnOnClick
iCurrent = 1
iPosition = iDOW
'dDate = DateValue(dDate)
do while iCurrent <= iDIM
	Rem -If we're at the begginning of a row then write tr
	Rem -If we're at the endof a row then write /tr
	if iPosition > 7 then
		Response.Write	"        </tr>" & vbNewLine
		iPosition = 1
	end if
	if iPosition = 1 then
		Response.Write	"        <tr>" & vbNewLine
	end if
	dCell = DateSerial(Year(dDate), Month(dDate), iCurrent)
	Rem -Get the current date in string Format (YYYYMMDD)
	strReturnFunc = "returnDate('" & Year(dDate) & doublenum(Month(dDate)) & doublenum(iCurrent) & "');"
	Rem -if Cell contains todays Date then highlight
	if dCell = dDate then 'and dDate < strForumTimeAdjust then
		strClass = ""
		blnOnClick = true
		Rem -if we are in the past then if history is 'off' the show cell disabled
	elseif dCell <= strForumTimeAdjust then
		if Request("History") = "on" then
			strClass = ""
			blnOnClick = false
		else
			strClass = "spnMessageText"
			blnOnClick = true
		end if
	Rem -else must be in the future
	else
		strClass = ""
		blnOnClick = false
	end if
	Call WriteDayOfMonth(strReturnFunc, strClass, iCurrent, blnOnClick)
	Rem -Increment variables
	iCurrent = iCurrent + 1
	iPosition = iPosition + 1
loop
Rem -Write spacer cells at end of last row if month doesn't end on a Saturday.
iPosition = iPosition - 1
if iPosition < 7 then
	Response.Write	"          <td colspan=""" & 7 - iPosition & """ bgcolor=""#bbbbbb""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>&nbsp;<br /><br /></font></td>" & vbNewLine & _
			"        </tr>" & vbNewLine
end if
Response.Write	"      </table>" & vbNewLine & _
		"      <table width=""275"" border=""0"" cellspacing=""0"" bgcolor=""" & strPageBGColor & """>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""center""><input type=""button"" value=""Clear DOB"" onclick=""" & strReturnFuncEmpty & """></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td align=""left""><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Birthdate Selection/Removal:</b><br /><li>Select Month and Year and press GO,<br />&nbsp;&nbsp;&nbsp;&nbsp;then click on the date.<br /><li>Click on ClearDOB to remove the Birthdate.</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine & _
		"    </td>" & vbNewLine & _
		"  </tr>" & vbNewLine & _
		"</table>" & vbNewLine & _
		"</form>" & vbNewLine & _
		"</body>" & vbNewLine & _
		"</html>" & vbNewline
%>