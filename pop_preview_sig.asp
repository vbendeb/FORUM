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
<!--#INCLUDE FILE="inc_header_short.asp"-->
<!--#INCLUDE file="inc_func_member.asp" -->
<%
Response.Write	"      <script language=""JavaScript"" type=""text/javascript"">" & vbNewLine & _
		"      function submitPreview()" & vbNewLine & _
		"      {" & vbNewLine & _
		"      document.previewSig.sig.value = window.opener.document.Form1.Sig.value;" & vbNewLine & _
		"      document.previewSig.submit()" & vbNewLine & _
		"      }" & vbNewLine & _
		"      </script>" & vbNewLine
if request("mode") = "" then
	Response.Write	"      <form action=""pop_preview_sig.asp"" method=""post"" name=""previewSig"">" & vbNewLine & _
			"      <input type=""hidden"" name=""sig"" value="""">" & vbNewLine & _
			"      <input type=""hidden"" name=""mode"" value=""display"">" & vbNewLine & _
			"      </form>" & vbNewLine & _
			"      <script language=""JavaScript"" type=""text/javascript"">submitPreview();</script>" & vbNewLine
else
	strSigPreview = trim(request.form("sig"))
	if strSigPreview = "" or IsNull(strSigPreview) then
		if strAllowForumCode = "1" then
			strSigPreview = "[center][b]< There is no text to preview ! >[/b][/center]"
		else
			strSigPreview = "<center><b>< There is no text to preview ! ></b></center>"
		end if
	end if
	Response.Write	"      <table border=""0"" width=""100%"" height=""80%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" width=""100%"" height=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td align=""center"" bgcolor=""" & strHeadCellColor & """ height=""20""><b><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strHeadFontColor & """>Signature Preview</font></b></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strForumCellColor & """ valign=""bottom""><hr noshade size=""" & strFooterFontSize & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """><span class=""spnMessageText"">" & formatStr(chkString(strSigPreview,"preview")) & "</span></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine
end if
WriteFooterShort
Response.End
%>