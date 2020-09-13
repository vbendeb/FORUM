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

Response.Write	"                      <script language=""Javascript"" type=""text/javascript"">" & vbNewLine & _
		"                      <!-- hide" & vbNewLine & _
        	"                      function insertsmilie(smilieface) {" & vbNewLine & _
		"                      		if (document.PostTopic.Message.createTextRange && document.PostTopic.Message.caretPos) {" & vbNewLine & _
		"                      			var caretPos = document.PostTopic.Message.caretPos;" & vbNewLine & _
		"                      			caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? smilieface + ' ' : smilieface;" & vbNewLine & _
		"                      			document.PostTopic.Message.focus();" & vbNewLine & _
		"                      		} else {" & vbNewLine & _
		"                      			document.PostTopic.Message.value+=smilieface;" & vbNewLine & _
		"                      			document.PostTopic.Message.focus();" & vbNewLine & _
		"                      		}" & vbNewLine & _
		"                      }" & vbNewLine & _
		"                      // -->" & vbNewLine & _
		"                      </script>" & vbNewLine & _
		"                      <table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""2"" align=""center"">" & vbNewLine & _
		"                        <tr align=""center"">" & vbNewLine & _
		"                          <td align=""center"" colspan=""4""><a name=""smilies""></a><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """><b>Smilies</b></font></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                        <tr align=""center"" valign=""middle"">" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmile,"Smile [:)]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:D]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileBig,"Big Smile [:D]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[8D]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileCool,"Cool [8D]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:I]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileBlush,"Blush [:I]","") & "</a></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                        <tr align=""center"" valign=""middle"">" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:p]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileTongue,"Tongue [:P]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[}:)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileEvil,"Evil [):]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[;)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileWink,"Wink [;)]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:o)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileClown,"Clown [:o)]","") & "</a></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                        <tr align=""center"" valign=""middle"">" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[B)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileBlackeye,"Black Eye [B)]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[8]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmile8ball,"Eight Ball [8]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:(]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileSad,"Frown [:(]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[8)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileShy,"Shy [8)]","") & "</a></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                        <tr align=""center"" valign=""middle"">" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:0]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileShock,"Shocked [:0]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:(!]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileAngry,"Angry [:(!]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[xx(]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileDead,"Dead [xx(]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[|)]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileSleepy,"Sleepy [|)]","") & "</a></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                        <tr align=""center"" valign=""middle"">" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[:X]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileKisses,"Kisses [:X]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[^]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileApprove,"Approve [^]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[V]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileDisapprove,"Disapprove [V]","") & "</a></td>" & vbNewLine & _
		"                          <td><a href=""Javascript:insertsmilie('[?]')"" tabindex=""-1"">" & getCurrentIcon(strIconSmileQuestion,"Question [?]","") & "</a></td>" & vbNewLine & _
		"                        </tr>" & vbNewLine & _
		"                      </table>" & vbNewLine
%>