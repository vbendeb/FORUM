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
if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

select case Request.QueryString("mode")
	case "AIM"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_AIM "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsAIM = my_Conn.execute(strSql)
		
		strProfileName = chkString(rsAIM("M_NAME"), "display")

		Response.Write	"      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>" & strProfileName & "'s AIM Options</font></p>" & vbNewLine & _
				"      " & strParagraphFormat1 & "<b>NOTE:</b> You must have AOL Instant Messenger installed in order for these functions to work properly.</font></p>" & vbNewLine & _
				"      <table border=""0"" width=""75%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""aim:goIM?screenname=" & rsAIM("M_AIM") & """ alt=""Opens a send message window to the user."">Send a Message</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""aim:goChat?ROOMname=" & rsAIM("M_AIM") & """>Open a chat room</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""center"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """><a href=""aim:addBuddy?screenname=" & rsAIM("M_AIM") & """>Add to buddy list</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine
		rsAIM.close
		set rsAIM = nothing
	case "ICQ"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_ICQ "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsICQ = my_Conn.execute(strSql)

		strProfileName = chkString(rsICQ("M_NAME"), "display")

		Response.Write	"      <p><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """>Send an ICQ Message</font></p>" & vbNewLine & _
				"      <form action=""http://wwp.icq.com/scripts/WWPMsg.dll"" method=""post"">" & vbNewLine & _
				"      <input type=""hidden"" name=""subject"" value=""" & strForumTitle & """>" & vbNewLine & _
				"      <input type=""hidden"" name=""to"" value=""" & rsICQ("M_ICQ") & """>" & vbNewLine & _
				"      <table border=""0"" width=""75%"" cellspacing=""0"" cellpadding=""0"">" & vbNewLine & _
				"        <tr>" & vbNewLine & _
				"          <td bgcolor=""" & strPopUpBorderColor & """>" & vbNewLine & _
				"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""1"">" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send to Name:&nbsp;</font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & strProfileName & "</font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Send to ICQ:&nbsp;</font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & rsICQ("M_ICQ") & "&img=5|18|18","","align=""absmiddle""") & rsICQ("M_ICQ") & "</font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your Name:&nbsp;</font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""from"" size=""20"" maxlength=""40"" onfocus=""this.select()""></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _ 
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Your E-mail:&nbsp;</font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><input type=""text"" name=""fromemail"" size=""20"" maxlength=""40"" onfocus=""this.select()""></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr valign=""top"">" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ align=""right"" nowrap><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>Message:&nbsp;</font></td>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """><textarea name=""body"" rows=""4"" cols=""25"" wrap=""Virtual""></textarea></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgColor=""" & strPopUpTableColor & """ Colspan=""2"" align=""center""><input type=""submit"" value=""Send""></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"            </table>" & vbNewLine & _
				"          </td>" & vbNewLine & _
				"        </tr>" & vbNewLine & _
				"      </table>" & vbNewLine & _
				"      </form>" & vbNewLine
		rsICQ.close
		set rsICQ = nothing
	case "MSN"
		'## Forum_SQL
		strSql = "SELECT MEMBER_ID, M_NAME, M_MSN "
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " WHERE MEMBER_ID = " & intMemberID

		Set rsMSN = my_Conn.execute(strSql)

		strProfileName = chkString(rsMSN("M_NAME"), "display")

		parts = split(rsMSN("M_MSN"),"@")
		strtag1 = parts(0)
		partss = split(parts(1),".")
		strtag2 = partss(0)
		strtag3 = partss(1)

		Response.Write	"      <script language=""javascript"" type=""text/javascript"">" & vbNewLine & _
				"              function MSNjs() {" & vbNewLine & _
				"              var tag1 = '" & strtag1 & "';" & vbNewLine & _
				"              var tag2 = '" & strtag2 & "';" & vbNewLine & _
				"              var tag3 = '" & strtag3 & "';" & vbNewLine & _
				"              document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & vbNewLine & _
				"      </script>" & vbNewLine

		Response.Write	"      " & strParagraphFormat1 & "<b>" & strProfileName & "'s MSN Messenger address:</b><br /><br /><br /><script language=""javascript"" type=""text/javascript"">MSNjs()</script></p><br /><br />" & vbNewLine
		rsMSN.close
		set rsMSN = nothing
end select
if (not(strUseExtendedProfile) and InStr(Request.ServerVariables("HTTP_REFERER"), "pop_profile.asp") <> 0) then Response.Write("      <p align=""center""><font size=""" & strDefaultFontSize & """><a href=""JavaScript:history.go(-1)"">Return to " & strProfileName & "'s Profile</a></font></p>" & vbNewLine)
WriteFooterShort
Response.End
%>