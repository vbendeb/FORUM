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

'## Set the default config-values
strDummy = SetConfigValue(1,"STRVERSION", strNewVersion)
strDummy = SetConfigValue(1,"STRFORUMTITLE","Snitz Forums 2000")
strDummy = SetConfigValue(1,"STRCOPYRIGHT","2000-02 Snitz Communications")
strDummy = SetConfigValue(1,"STRTITLEIMAGE","logo_snitz_forums_2000.gif")
strDummy = SetConfigValue(1,"STRHOMEURL","../")
strDummy = SetConfigValue(1,"STRFORUMURL","http://www.yourdomain.com/forum/")
strDummy = SetConfigValue(1,"STRAUTHTYPE","db")
strDummy = SetConfigValue(1,"STRSETCOOKIETOFORUM","1")
strDummy = SetConfigValue(1,"STREMAIL","0")
strDummy = SetConfigValue(1,"STRUNIQUEEMAIL","1")
strDummy = SetConfigValue(1,"STRMAILMODE","cdonts")
strDummy = SetConfigValue(1,"STRMAILSERVER","your.mailserver.com")
strDummy = SetConfigValue(1,"STRSENDER","your.email@yourserver.com")
strDummy = SetConfigValue(1,"STRDATETYPE","mdy")
strDummy = SetConfigValue(1,"STRTIMETYPE","24")
strDummy = SetConfigValue(1,"STRTIMEADJUSTLOCATION","0")
strDummy = SetConfigValue(1,"STRTIMEADJUST","0")
strDummy = SetConfigValue(1,"STRMOVETOPICMODE","1")
strDummy = SetConfigValue(1,"STRPRIVATEFORUMS","0")
strDummy = SetConfigValue(1,"STRSHOWMODERATORS","1")
strDummy = SetConfigValue(1,"STRSHOWRANK","0")
strDummy = SetConfigValue(1,"STRHIDEEMAIL","0")
strDummy = SetConfigValue(1,"STRIPLOGGING","1")
strDummy = SetConfigValue(1,"STRALLOWFORUMCODE","1")
strDummy = SetConfigValue(1,"STRIMGINPOSTS","0")
strDummy = SetConfigValue(1,"STRALLOWHTML","0")
strDummy = SetConfigValue(1,"STRSECUREADMIN","1")
strDummy = SetConfigValue(1,"STRNOCOOKIES","0")
strDummy = SetConfigValue(1,"STREDITEDBYDATE", "1")
strDummy = SetConfigValue(1,"STRHOTTOPIC","1")
strDummy = SetConfigValue(1,"INTHOTTOPICNUM","20")
strDummy = SetConfigValue(1,"STRHOMEPAGE","1")
strDummy = SetConfigValue(1,"STRAIM","1")
strDummy = SetConfigValue(1,"STRICQ","1")
strDummy = SetConfigValue(1,"STRMSN","1")
strDummy = SetConfigValue(1,"STRYAHOO","1")
strDummy = SetConfigValue(1,"STRICONS","1")
strDummy = SetConfigValue(1,"STRGFXBUTTONS","1")
strDummy = SetConfigValue(1,"STRBADWORDFILTER","1")
strDummy = SetConfigValue(1,"STRBADWORDS","fuck|wank|shit|pussy|cunt")
strDummy = SetConfigValue(1,"STRUSERNAMEFILTER","0")
strDummy = SetConfigValue(1,"STRDEFAULTFONTFACE","Verdana, Arial, Helvetica")
strDummy = SetConfigValue(1,"STRDEFAULTFONTSIZE","2")
strDummy = SetConfigValue(1,"STRHEADERFONTSIZE","4")
strDummy = SetConfigValue(1,"STRFOOTERFONTSIZE","1")
strDummy = SetConfigValue(1,"STRPAGEBGCOLOR","white")
strDummy = SetConfigValue(1,"STRDEFAULTFONTCOLOR","midnightblue")
strDummy = SetConfigValue(1,"STRLINKCOLOR","darkblue")
strDummy = SetConfigValue(1,"STRLINKTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRVISITEDLINKCOLOR","blue")
strDummy = SetConfigValue(1,"STRVISITEDTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRACTIVELINKCOLOR","red")
strDummy = SetConfigValue(1,"STRHOVERFONTCOLOR","red")
strDummy = SetConfigValue(1,"STRHOVERTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRHEADCELLCOLOR","midnightblue")
strDummy = SetConfigValue(1,"STRHEADFONTCOLOR","mintcream")
strDummy = SetConfigValue(1,"STRCATEGORYCELLCOLOR","slateblue")
strDummy = SetConfigValue(1,"STRCATEGORYFONTCOLOR","mintcream")
strDummy = SetConfigValue(1,"STRFORUMFIRSTCELLCOLOR","whitesmoke")
strDummy = SetConfigValue(1,"STRFORUMCELLCOLOR","whitesmoke")
strDummy = SetConfigValue(1,"STRALTFORUMCELLCOLOR","gainsboro")
strDummy = SetConfigValue(1,"STRFORUMFONTCOLOR","midnightblue")
strDummy = SetConfigValue(1,"STRFORUMLINKCOLOR","darkblue")
strDummy = SetConfigValue(1,"STRFORUMLINKTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRFORUMVISITEDLINKCOLOR","blue")
strDummy = SetConfigValue(1,"STRFORUMVISITEDTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRFORUMACTIVELINKCOLOR","red")
strDummy = SetConfigValue(1,"STRFORUMACTIVETEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRFORUMHOVERFONTCOLOR","red")
strDummy = SetConfigValue(1,"STRFORUMHOVERTEXTDECORATION","underline")
strDummy = SetConfigValue(1,"STRTABLEBORDERCOLOR","black")
strDummy = SetConfigValue(1,"STRPOPUPTABLECOLOR","lightsteelblue")
strDummy = SetConfigValue(1,"STRPOPUPBORDERCOLOR","black")
strDummy = SetConfigValue(1,"STRNEWFONTCOLOR","blue")
strDummy = SetConfigValue(1,"STRHILITEFONTCOLOR","red")
strDummy = SetConfigValue(1,"STRSEARCHHILITECOLOR","yellow")
strDummy = SetConfigValue(1,"STRTOPICWIDTHLEFT","100")
strDummy = SetConfigValue(1,"STRTOPICWIDTHRIGHT","100%")
strDummy = SetConfigValue(1,"STRTOPICNOWRAPLEFT","1")
strDummy = SetConfigValue(1,"STRTOPICNOWRAPRIGHT","0")
strDummy = SetConfigValue(1,"STRRANKADMIN","Administrator")
strDummy = SetConfigValue(1,"STRRANKMOD","Moderator")
strDummy = SetConfigValue(1,"STRRANKLEVEL0","Starting Member")
strDummy = SetConfigValue(1,"STRRANKLEVEL1","New Member")
strDummy = SetConfigValue(1,"STRRANKLEVEL2","Junior Member")
strDummy = SetConfigValue(1,"STRRANKLEVEL3","Average Member")
strDummy = SetConfigValue(1,"STRRANKLEVEL4","Senior Member")
strDummy = SetConfigValue(1,"STRRANKLEVEL5","Advanced Member")
strDummy = SetConfigValue(1,"STRRANKCOLORADMIN","gold")
strDummy = SetConfigValue(1,"STRRANKCOLORMOD","silver")
strDummy = SetConfigValue(1,"STRRANKCOLOR0","bronze")
strDummy = SetConfigValue(1,"STRRANKCOLOR1","bronze")
strDummy = SetConfigValue(1,"STRRANKCOLOR2","bronze")
strDummy = SetConfigValue(1,"STRRANKCOLOR3","bronze")
strDummy = SetConfigValue(1,"STRRANKCOLOR4","bronze")
strDummy = SetConfigValue(1,"STRRANKCOLOR5","bronze")
strDummy = SetConfigValue(1,"INTRANKLEVEL0","0")
strDummy = SetConfigValue(1,"INTRANKLEVEL1","50")
strDummy = SetConfigValue(1,"INTRANKLEVEL2","100")
strDummy = SetConfigValue(1,"INTRANKLEVEL3","500")
strDummy = SetConfigValue(1,"INTRANKLEVEL4","1000")
strDummy = SetConfigValue(1,"INTRANKLEVEL5","2000")
strDummy = SetConfigValue(1,"STRSIGNATURES","1")
strDummy = SetConfigValue(1,"STRDSIGNATURES", "0")
strDummy = SetConfigValue(1,"STRSHOWSTATISTICS","1")
strDummy = SetConfigValue(1,"STRSHOWIMAGEPOWEREDBY","1")
strDummy = SetConfigValue(1,"STRLOGONFORMAIL","1")
strDummy = SetConfigValue(1,"STRSHOWPAGING","1")
strDummy = SetConfigValue(1,"STRSHOWTOPICNAV","1")
strDummy = SetConfigValue(1,"STRPAGESIZE","15")
strDummy = SetConfigValue(1,"STRPAGENUMBERSIZE","10")
strDummy = SetConfigValue(1,"STRFULLNAME","1")
strDummy = SetConfigValue(1,"STRPICTURE","0")
strDummy = SetConfigValue(1,"STRSEX","0")
strDummy = SetConfigValue(1,"STRCITY","0")
strDummy = SetConfigValue(1,"STRSTATE","0")
strDummy = SetConfigValue(1,"STRAGE","0")
strDummy = SetConfigValue(1,"STRAGEDOB","0")
strDummy = SetConfigValue(1,"STRCOUNTRY","1")
strDummy = SetConfigValue(1,"STROCCUPATION","0")
strDummy = SetConfigValue(1,"STRHOMEPAGE","1")
strDummy = SetConfigValue(1,"STRFAVLINKS","1")
strDummy = SetConfigValue(1,"STRBIO","0")
strDummy = SetConfigValue(1,"STRHOBBIES","0")
strDummy = SetConfigValue(1,"STRLNEWS","0")
strDummy = SetConfigValue(1,"STRQUOTE","0")
strDummy = SetConfigValue(1,"STRMARSTATUS","0")
strDummy = SetConfigValue(1,"STRRECENTTOPICS","1")
strDummy = SetConfigValue(1,"STRNTGROUPS","0")
strDummy = SetConfigValue(1,"STRAUTOLOGON","0")
strDummy = SetConfigValue(1,"STRMOVENOTIFY","0")
strDummy = SetConfigValue(1,"STRSUBSCRIPTION", "0")
strDummy = SetConfigValue(1,"STRMODERATION", "0")
strDummy = SetConfigValue(1,"STRARCHIVESTATE", "1")
strDummy = SetConfigValue(1,"STRFLOODCHECK", "1")
strDummy = SetConfigValue(1,"STRFLOODCHECKTIME", "-60")
strDummy = SetConfigValue(1,"STREMAILVAL", "0")
strDummy = SetConfigValue(1,"STRPAGEBGIMAGEURL", "")
strDummy = SetConfigValue(1,"STRIMAGEURL", "")
strDummy = SetConfigValue(1,"STRJUMPLASTPOST", "0")
strDummy = SetConfigValue(1,"STRSTICKYTOPIC", "0")
strDummy = SetConfigValue(1,"STRSHOWSENDTOFRIEND", "1")
strDummy = SetConfigValue(1,"STRSHOWPRINTERFRIENDLY", "1")
strDummy = SetConfigValue(1,"STRPROHIBITNEWMEMBERS", "0")
strDummy = SetConfigValue(1,"STRREQUIREREG", "0")
strDummy = SetConfigValue(1,"STRRESTRICTREG", "0")
strDummy = SetConfigValue(1,"STRGROUPCATEGORIES", "0")
strDummy = SetConfigValue(1,"STRSHOWTIMER", "0")
strDummy = SetConfigValue(1,"STRTIMERPHRASE","This page was generated in [TIMER] seconds.")
strDummy = SetConfigValue(1,"STRSHOWFORMATBUTTONS","1")
strDummy = SetConfigValue(1,"STRSHOWSMILIESTABLE","1")
strDummy = SetConfigValue(1,"STRSHOWQUICKREPLY","0")
%>