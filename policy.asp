<%@CODEPAGE=1251%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1251">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title></title>
</head>

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
Response.Write	"      <table width=""100%"" border=""0"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """>" & vbNewLine & _
		"          " & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">All Forums</a><br />" & vbNewLine & _
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;������� ����������� � �����������</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if strProhibitNewMembers <> "1" then
	Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>������� ����������� � ������� ����������� " & strForumTitle & "</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
			"                <p><U>������� �����������</U><BR> " & _
			"<font color=""" & strHiLiteFontColor & """>��������: ��� �� ������� ������������ ���� ����� � �������� ������ � ������ ��� ����������� E-mail ������. ���� �� �� �������� ���������������� ���������, ��������� ����������� �������, �� ���� ��� � ���������� ������� E-mail.  �������� ��� �������� ������� ������ ��� ������ ��� ��������� �����������.</p></font>" & vbNewLine & _
			"1. ����� ���������� ������ � ������� ������ &quot;�����������&quot; �� ��������� �������� ��� ������ ���������� �� MOCT.<BR> " & _
			"2. ��������� ���������� ���� ����������� �� ������ (��� ����� ������ ��������� ����!).<BR> " & _
			"3. � ���������� ������������� �� ��� E-mail ����� ����� ������� ��������������� ��������� � ����������� ������� �� �����.<BR> " & _
			"4. ������� (����) �� ������ � ��������������� ��������� ������������ ���� �����������.<BR> " & _
			"5. ��� �4 ������������ ������������ ������ E-mail ������. ���� ����� ������������, ������� ����������� �� �2.<BR> " & _
			" <BR> " & _
			"                <p><U>�������� ��� �����������</U><BR> " & _
			"����� ������ �������� - &quot;������������� �����������&quot;. ��� ����� ��������� � ��������� �������:<BR> " & _
			"1. �������� ������������� E-Mail ������ (��������).<BR> " & _
			"2. C�������� ��������� ��������� �� ��������� ����� �� ���� �������� ��������� (���������� �����).<BR> " & _
			"3. C�������� ��������� ��������� �� ��������� ����� �� ���� ������������ ���������.<BR> " & _
			"� ����� �� ���� �������, ����� ��� &quot;��������&quot; � ���������� ����������� ��� ���� �� ����� ����� ���������. � ���� ������, ������� E-Mail �� ����� <A href='mailto:administrator@moct.org'>administrator@moct.org</A> � ������������ ��������� ���� � E-Mail ������.<BR> " & _
			" <BR> " & _
			"                <p><U>������� �����������</U><BR> " & _
			"������������� ������ �� ����� ��������� �������������� ���������� ��������� � ������� �� � ��������� �������� " & _
			"��� �� ����-�� ��� ����� ��� �� ����������� ��� ���� ���������. ��� �� �����, �� ����� ������� �� ��� �� ��������� ������������ � ������� ��� ������ ������. " & _
			"�������� �� ������ &quot;�����������&quot; �� ��������� ���� �������� � ��������� " & _
			"� ���������� �� ���� ������ ��������������� �� ���� �������� �� ������ � �� ���������� �� ����������. " & _
			"�� ����� �������� �� ����������� ������� ���������� ���������� � �� �������� ������� ���������� �����. " & _
			"�� ����� ���������� �� ����������� ������� ��������� �����������, ���������������, ����������������� ��� ������������ ���������. </p>" & vbNewLine & _
			"                <p><U>������������ ����������</U><BR> " & _
			"<font color=""" & strHiLiteFontColor & """>������������ ���������� ������������ � ����� ������������ ������� ���������� � ����������� ��� ���� ���������� ����� � ������.  ���� ������� ������ - ������ ����������� ��� � �� ������� ����� ���� �����, ��� ���������� � ������ ���������� ��� �����������.  ���� ������������ ���������� �� ����� � �� ����� ������������ �� � ����� ������ �����!  � ����� ������, ��������� ������������ ����� ��������� �� ������ ���������� � �������� ������� � ������.  �� ������ ��������������� ����� ��� ����� ��������, ���� ������� ��������� � ������ ������ � ��������� IP ����� ������� ������������, �� �������� ����� ����� ����� �� ����������, ��������, ����� ��� ����������� ������� ���� ����� �����������. ��������� ����� �������� ��� �������� ���������� � ���� ��������� ���� ����� � ������ ��� ��������� ������ � �� ������������ �������. � ������ ����� ������, �� ������ ������� ������ �� E-Mail ����� " & _
			"  <span class=""spnMessageText""><a href=""mailto:" & strSender & """>" & strSender & "</a></span>.</font>" & _
			"</ol>" & _
			"</p>" & vbNewLine & _
			"                <p>���� �� �������� � ��������� ������ ����������� ����, ���������� ������� ������ &quot;������� � �����������&quot;, � ��������� ������ ������� ������ &quot;������&quot;.</p>" & vbNewLine & _
			"                <hr size=""1"">" & vbNewLine & _
			"                  <table align=""center"" border=""0"">" & vbNewLine & _
			"                    <tbody>" & vbNewLine & _
			"                      <tr>" & vbNewLine & _
			"                        <td>" & vbNewLine & _
			"                        <form action=""register.asp?mode=Register"" id=""form1"" method=""post"" name=""form1"">" & vbNewLine & _
			"                        <input name=""Refer"" type=""hidden"" value=""" & Request.ServerVariables("HTTP_REFERER") & """>" & vbNewLine & _
			"                        <input name=""Submit"" type=""Submit"" value=""������� � �����������"">" & vbNewLine & _
			"                        </form>" & vbNewLine & _
			"                        </td>" & vbNewLine & _
			"                        <td>" & vbNewLine & _
			"                        <form action=""JavaScript:history.go(-1)"" id=""form2"" method=""post"" name=""form2"">" & vbNewLine & _
			"                        <input name=""Submit"" type=""Submit"" value=""������"">" & vbNewLine & _
			"                        </form>" & vbNewLine & _
			"                        </td>" & vbNewLine & _
			"                      </tr>" & vbNewLine & _
			"                    </tbody>" & vbNewLine & _
			"                  </table>" & vbNewLine & _
			"                <hr size=""1"">" & vbNewLine & _
			"                <p>�� ���� �������� ������ " & _
			"�� ������ ��������� � ���� �� ������: " & _
			"<span class=""spnMessageText""><a href=""mailto:" & strSender & """>" & strSender & "</a></span></p>" & vbNewLine & _
			"                </font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"            </table>" & vbNewLine & _
			"          </td>" & vbNewLine & _
			"        </tr>" & vbNewLine & _
			"      </table>" & vbNewLine & _
			"      <br />" & vbNewLine
else
	Response.Write	"    <br /><p align=""center""><font face=""" & strDefaultFontFace & """ size=""" & strHeaderFontSize & """ color=""" & strHiLiteFontColor & """>Sorry, we are not accepting any new Members at this time.</font></p>" & vbNewLine & _
			"    <meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & vbNewLine & _
 			"    " & strParagraphFormat1 & "<a href=""default.asp"">" & strBackToForum & "</font></a></p><br />" & vbNewLine
end if
WriteFooter
Response.End
%>