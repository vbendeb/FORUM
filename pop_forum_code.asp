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
Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
		"        <tr>" & vbNewLine & _
		"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strCategoryCellColor & """><a name=""format""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>How to format text with Bold, Italic, Quote, etc...</b></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"              <tr>" & vbNewLine & _
		"                <td bgcolor=""" & strForumCellColor & """>" & vbNewLine & _
		"                <p><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                There are several Forum Codes you may use to change the appearance " & vbNewLine & _
		"                of your text.&nbsp; Following is the list of codes currently available:</p>" & vbNewLine & _
		"                <blockquote>" & vbNewLine & _
		"                <p><b>Bold:</b> Enclose your text with [b] and [/b] .&nbsp; <i>Example:</i> This is <b>[b]</b>bold<b>[/b]</b> text. = This is <b>bold</b> text.</p>" & vbNewLine & _
		"                <p><i>Italic:</i> Enclose your text with [i] and [/i] .&nbsp; <i>Example:</i> This is <b>[i]</b>italic<b>[/i]</b> text. = This is <i>italic</i> text.</p>" & vbNewLine & _
		"                <p><u>Underline:</u> Enclose your text with [u] and [/u]. <i>Example:</i> This is <b>[u]</b>underline<b>[/u]</b> text. =  This is <u>underline</u> text.</p>" & vbNewLine & _
		"                <p><b>Aligning Text Left:</b><br />" & vbNewLine & _
		"                Enclose your text with [left] and [/left]" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p><b>Aligning Text Center:</b><br />" & vbNewLine & _
		"                Enclose your text with [center] and [/center]" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p><b>Aligning Text Right:</b><br />" & vbNewLine & _
		"                Enclose your text with [right] and [/right]" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p><b>Striking Text:</b><br />" & vbNewLine & _
		"                Enclose your text with [s] and [/s]<br />" & vbNewLine & _
		"                <i>Example:</i> <b>[s]</b>mistake<b>[/s]</b> = <s>mistake</s>" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p><b>Horizontal Rule:</b><br />" & vbNewLine & _
		"                Place a horizontal line in your post with [hr]<br />" & vbNewLine & _
		"                <i>Example:</i> <b>[hr]</b> = <hr noshade size=""1"">" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p>&nbsp; </p>" & vbNewLine & _
		"                <p><b>Font Colors:</b><br />" & vbNewLine & _
		"                Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>] <br />" & vbNewLine & _
		"                <i>Example:</i> <b>[red]</b>Text<b>[/red]</b> = <font color=""red"">Text</font id=""red""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[blue]</b>Text<b>[/blue]</b> = <font color=""blue"">Text</font id=""blue""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[pink]</b>Text<b>[/pink]</b> = <font color=""pink"">Text</font id=""pink""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[brown]</b>Text<b>[/brown]</b> = <font color=""brown"">Text</font id=""brown""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[black]</b>Text<b>[/black]</b> = <font color=""black"">Text</font id=""black""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[orange]</b>Text<b>[/orange]</b> = <font color=""orange"">Text</font id=""orange""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[violet]</b>Text<b>[/violet]</b> = <font color=""violet"">Text</font id=""violet""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[yellow]</b>Text<b>[/yellow]</b> = <font color=""yellow"">Text</font id=""yellow""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[green]</b>Text<b>[/green]</b> = <font color=""green"">Text</font id=""green""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[gold]</b>Text<b>[/gold]</b> = <font color=""gold"">Text</font id=""gold""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[white]</b>Text<b>[/white]</b> = <font color=""white"">Text</font id=""white""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[purple]</b>Text<b>[/purple]</b> = <font color=""purple"">Text</font id=""purple"">" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p>&nbsp; </p>" & vbNewLine & _
		"                <p><b>Headings:</b><br />" & vbNewLine & _
		"                Enclose your text with [h<i>number</i>] and [/h<i>n</i>]<br />" & vbNewLine & _
		"                  <table border=""0"">" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h1]</b>Text<b>[/h1]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h1>Text</h1>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h2]</b>Text<b>[/h2]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h2>Text</h2>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h3]</b>Text<b>[/h3]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h3>Text</h3>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h4]</b>Text<b>[/h4]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h4>Text</h4>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h5]</b>Text<b>[/h5]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h5>Text</h5>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                    <tr>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <i>Example:</i> <b>[h6]</b>Text<b>[/h6]</b> =" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                      <td><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
		"                      <h6>Text</h6>" & vbNewLine & _
		"                      </font></td>" & vbNewLine & _
		"                    </tr>" & vbNewLine & _
		"                  </table>" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p>&nbsp; </p>" & vbNewLine & _
		"                <p><b>Font Sizes:</b><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=1]</b>Text<b>[/size=1]</b> = <font size=""1"">Text</font id=""size1""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=2]</b>Text<b>[/size=2]</b> = <font size=""2"">Text</font id=""size2""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=3]</b>Text<b>[/size=3]</b> = <font size=""3"">Text</font id=""size3""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=4]</b>Text<b>[/size=4]</b> = <font size=""4"">Text</font id=""size4""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=5]</b>Text<b>[/size=5]</b> = <font size=""5"">Text</font id=""size5""><br />" & vbNewLine & _
		"                <i>Example:</i> <b>[size=6]</b>Text<b>[/size=6]</b> = <font size=""6"">Text</font id=""size6"">" & vbNewLine & _
		"                </p>" & vbNewLine & _
		"                <p>&nbsp; </p>" & vbNewLine & _
		"                <p><b>Bulleted List:</b> <b>[list]</b> and <b>[/list]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & vbNewLine & _
		"                <p><b>Ordered Alpha List:</b> <b>[list=a]</b> and <b>[/list=a]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & vbNewLine & _
		"                <p><b>Ordered Number List:</b> <b>[list=1]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & vbNewLine & _
		"                <p><b>Code:</b> Enclose your text with <b>[code]</b> and <b>[/code]</b>.</p>" & vbNewLine & _
		"                <p><b>Quote:</b> Enclose your text with <b>[quote]</b> and <b>[/quote]</b>.</p>" & vbNewLine
if (strIMGInPosts = "1") then
	Response.Write	"              <p><b>Images:</b> Enclose the address with one of the following:<ul><li><b>[img]</b> and <b>[/img]</b></li>" & vbNewLine & _
			"              <li><b>[img=right]</b> and <b>[/img=right]</b></li>" & vbNewLine & _
			"              <li><b>[img=left]</b> and <b>[/img=left]</b></li></ul></p>" & vbNewLine
end if
Response.Write	"                </blockquote></font></td>" & vbNewLine & _
		"              </tr>" & vbNewLine & _
		"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
WriteFooterShort
Response.End
%>