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
		"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine
select case Request.QueryString("mode")
	case "system"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""strConnString""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>How do I configure the strConnString?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """>" & vbNewLine & _
				"                <li><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """><b>DSN:</b></font><br />" & vbNewLine & _
				"                <font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>snitz_forum</font></li>" & vbNewLine & _
				"                <li><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """><b>MS Access DSN-less:</b></font><br />" & vbNewLine & _
				"                <font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>strConnString = &quot;DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\www\snitz.com\db\snitz_forum.mdb&quot;</font></li>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""tableprefix""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's Table Name Prefix?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Table Name Prefix is used if you have multiple versions of the forum running in the same database. This way you can name the tables differently and still use one user to connect. (eg. FORUM_ and FORUM2_)</font>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""forumtitle""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's Forum Title?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Forum Title is the title that shows up in the upper right hand corner of the forum. It is also used in e-mails to show where the e-mail came from when posting replies are sent and when new users register." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""copyright""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's Forum Copyright?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This copyright statements location is basically saying that any topics or replies that are posted are copyrighted material of your organization. This copyright location also helps to copyright the images of your logo and any other material that may be posted on forum pages; however, it is understood by copyright statements in code and informational pages, that the forum code itself is still copyright &copy; 2000 Snitz Communications.<br /><br /><b><font color=""" & strHiLiteFontColor & """>NOTE:</b>  The &copy; will be included automatically.</font>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""titleimage""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's Title Image?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Use a relative path to point to the image you want to show up in the upper left-hand corner of your forum window.<br />" & vbNewLine & _
				"                <br />" & vbNewLine & _
				"                For example:<br />" & vbNewLine & _
				"                <b>bboard_snitz.gif</b><br />" & vbNewLine & _
				"                This points to the bboard_snitz.gif graphic in the same directory, whereas the following would point to the root of the web server and up into the base /images/ directory:<br />" & vbNewLine & _
				"                <b>../images/bboard_snitz.gif</b>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""homeurl""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's the Home URL?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                The Home URL is the base address for your website. An example would be:<br />" & vbNewLine & _
				"                <b>forum.snitz.com</b><br />" & vbNewLine & _
				"                <br />" & vbNewLine & _
				"                <font color=""" & strHiLiteFontColor & """>NOTE: Include the full path of the URL whether it begins with <b>http://</b> in front or a relative URL such as <b>../</b>.</font>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""forumurl""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What's the Forum URL?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                The Forum URL is the base address for your forum. An example would be:<br />" & vbNewLine & _
				"                <b>http://forum.snitz.com/forum</b>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""imagelocation""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is the Images Location?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Enter the location where your images are located.<br />" & vbNewLine & _
				"                If you have not moved the images from their default location, then just leave this field blank.<br /><br />" & vbNewLine & _
				"                But, if you have created an <b>images</b> directory in your <b>forum</b> directory then enter:<br /><br />" & vbNewLine & _
				"                <b>images/</b><br /><br />in the field." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""AuthType""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Authorization Type?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                You can either select DataBase or NT Domain authorization." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""SetCookieToForum""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Set Cookie To...</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                You can tell your forum to set it's cookie to either the forum, or the base website. You would set it to the forum if you were hosting multiple forums on the same server or the same domain, and they each had different user communities, otherwise you want this feature set to Website and NOT Forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""GfxButtons""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Graphic Buttons?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                By enabling this feature, the forums will use pictures/graphics instead of the default buttons for ""Submit"" and ""Reset"" etc..." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""PoweredBy""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Use Graphic for ""Powered By"" link?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Toggles between using a Graphic Powered By Link, or a Text Powered By Link.  Either way, you must have one or the other..." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ProhibitNewMembers""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Prohibit New Members?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Toggles between allowing or disallowing people to Register on your Forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""RequireReg""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Require Registration?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                When this option in set to <b>On</b>, only registered members who are logged in will be able to view your Forum.  Everyone else will be presented with a login screen." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""UserNameFilter""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>UserName Filter?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                When this option in set to <b>On</b>, the names (or names that contain words) that you specify in the UserName Filter configuration will not be available for user's to register with." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "features"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""secureadminmode""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Secure Admin Mode?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """  color=""" & strHiLiteFontColor & """>" & vbNewLine & _
				"                <b>WARNING: Only turn Secure Admin off if you absolutely need to. If this option is turned off, anyone can change your forum's configuration!</b>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""allownoncookies""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why would I want Non-Cookie Mode on?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                If your user base does not use cookies, then you would want to turn this function ""ON"". WARNING: all your admin functions will be visible to all users if this function is ""ON""." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""IPLogging""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is IP Logging?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                IP Logging will record in the database the IP address of the person who posted a new Topic or Reply. A moderator or administrator then could view the IP by clicking on an icon above the post in the topic." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""FloodCheck""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Flood Control?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                With Flood Control enabled, normal users will have to wait the specified amount of time between posts before they can post again." & vbNewLine & _
				"                <br /><br />Admins and Moderators are not affected by this limitation." & vbNewLine & _
				"                <br /><br />You can choose 30 seconds, 60 seconds, 90 seconds or 120 seconds." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""privateforums""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What are Private Forums?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Private Forums enable you to only allow certain members to see that the forum exists. If it's only password protected, everyone can see that it exists, however, they are prompted for a password to get in." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""groupcategories""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What are Group Categories?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Group Categories enable you to ""group"" Categories together into ""Groups"" to better organize how Categories are displayed on your forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Subscription""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Highest level of Subscription for?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allows you to set the Highest Level of Subscription that can be used on the Forum.  You will also need to set the individual level in each of your Categories and Forums." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""badwordfilter""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Bad Word Filter?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Screen out words you and your guests would find offensive.<br /><br />Bad Words can by configured via the Bad Word Configuration option in the Admin Options." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Moderation""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Allow Topic Moderation do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                When enabled, this feature allows the Administrator or the Moderator to ""Approve"", ""Hold"" or ""Delete"" a users post before it is shown to the public." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowModerator""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Moderators do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Basically, if this function is on, it shows the name of the moderator beside the forum that they moderate on the main default page. If it is off, however, visitors won't see who is moderating the forum they are posting in." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""MoveTopicMode""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why Restrict Moderators from Moving Posts?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This feature either allows or dis-allows a Moderator of one forum to move topics within their forum to someone else's forum where they do not have moderator rights." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""MoveNotify""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Can I notify the Author if his Topic is moved?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                If enabled, this feature automatically sends an e-mail to the topic author if it is moved." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ArchiveState""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What are Archive Functions?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This toggles whether the icons/links show up for the Archive Functions of this Forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""stats""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Detailed Statistics do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of detailed statistics (last visited date and time, last post, active topics, newest member) at the bottom of the forum." & vbNewLine & _
				"                When turned off, some statistics are displayed at the top of the page." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""JumpLastPost""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Jump To Last Post Link do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of a Jump To Last Post Link " & getCurrentIcon(strIconLastpost,"","align=""absmiddle""") & " icon on the Default page, Forum page and Active Topics page.  This link will take the user to the last post in that topic." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""showpaging""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Quick Paging do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Quick Paging is when you have a topic that is more than 1 page, a small graphic and the #'s will be show next to the topic title so you can go straight to page 2 or 3, etc..." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""pagenumbersize""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Pagenumbers per row for?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This is now only used for the Topic Paging, it limits the amount of pages shown in each row when a topic is more than one page long." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""StickyTopic""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Allow Sticky Topics do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the ability of an Admin or Moderator to ""Stick"" a post at the top of the Topics List.  While this Topic is ""Sticky"", it will remain at the top of the list." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""editedbydate""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What would Edited By on Date do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                When a post is edited, there is an appending to the end of the post that says when and by whom the post was edited. Turning this function off would make it so that the footer would not be placed on the end of the post." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowTopicNav""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Prev / Next Topic do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of previous topic " & getCurrentIcon(strIconGoLeft,"","align=""absmiddle""") & " and next topic " & getCurrentIcon(strIconGoRight,"","align=""absmiddle""") & " icons on the topics page." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowSendToFriend""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Send to a Friend Link do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of a Send Topic to a Friend Link that is shown when viewing a topic..  This link will allow a user to e-mail a topic to a friend.  E-mail functions must be on for this link to show up." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowPrinterFriendly""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Printer Friendly Link do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of a Printer Friendly link that is shown when viewing a topic.  This link will popup a window with the topic and any replies that are shown in a format that is easier to print." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""hottopics""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What are Hot Topics?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Hot Topics change the topic folder icon in the Forum view from a normal folder to a flaming folder to let people know that your minimum number of posts has been met to categorize this topic as one that is seeing a lot of action." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""pagesize""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Items per page for?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This is the maximum amount of items shown on each page. Once the amount of items on the page reaches this amount, a dropdown box will be shown where you can select other pages." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""AllowHTML""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why would I allow HTML?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                By allowing HTML you are opening up a whole big can of worms. You may wish to allow HTML in a controlled INTRANET environment,though. It is not recommended to be used on the INTERNET as anyone can post anything without your being able to screen it. IE Pornographic pictures, JavaScript that messes up your pages, etc..." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""AllowForumCode""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Enable Forum Code?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                By turning off Forum Code, you can allow users to mark up their posts with safe codes." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""imginposts""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why enable Images in Posts?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allows users to place images into their Posts. However, you should be aware that this feature would allow anyone to post ANY image in your forums. This may lead to broken links and potentially objectionable material being displayed." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""icons""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What do Icons do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow users to post smiley faces and other icons allowed by the Forums within the body of their posts!" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""signatures""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why enable Signatures?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allows users to set a ""Signature"" into their Posts. The same concerns mentioned for Images in Posts applies here as well." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""dsignatures""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Why enable Dynamic Signatures?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                First, you must have Signatures enabled to use Dynamic Signatures.  With Dynamic Signatures enabled, the users signature is not added to the post until it is viewed, so if a person changes their signature, that change will apply to all posts made by that user.  But, this will only apply to posts made while Dyanmic Signatures are enabled.  Any signature that is already in a post won't be updated." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowFormatButtons""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Show Format Buttons?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This turns off or on the Format Section on the screen where your users post new topics/reply to existing topics.<br /><br /><font color=""" & strHiLiteFontColor & """>Note:</font>&nbsp;You must also have Forum Code enabled on your forum to use this feature." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowSmiliesTable""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Show Smilies Table?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allows users to insert smilies in their posts by clicking on the smilie in a small table shown to them in the post screen.<br /><br /><font color=""" & strHiLiteFontColor & """>Note:</font>&nbsp;You must also have Icons enabled on your forum to use this feature." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowQuickReply""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Show Quick Reply?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allows users to reply to a topic via a reply box at the bottom of the page when viewing a topic." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""timer""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Show Timer do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off the display of the time it took (in seconds) to generate/display the current page.  This time is shown in the footer of every (non popup) page." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""timerphrase""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Timer Phrase?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This is what will display in the footer of every (non popup) page.  The phrase must contain the <b>[TIMER]</b> placeholder.  This is where the actual time will be in the phrase (it's dynamically inserted when the page is created)." & vbNewLine & _
				"                <br /><br /><b><font color=""" & strHiLiteFontColor & """>Show Timer must be enabled for this to be used.</font></b>" & vbNewLine & _
				"                <br /><br />The default is:  <b>This page was generated in [TIMER] seconds.</b>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "members"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""FullName""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Fullname For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Full Name (First Name and Last Name), to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Picture""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Picture For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter a link to a Picture of themselves, to be viewed in their profile.<br /><br />As Admin, you should review the picture in each user's profile from time to time to be sure that the Picture linked to is appropriate for your Forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""RecentTopics""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Recent Topics For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                When Recent Topics is enabled, a list of the last 10 Topics posted to by a user will be shown in their Profile.<br /><br />This includes New Topics and replies to existing topics." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Sex""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Sex For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Sex (either Male or Female), to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Age""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Age For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their age, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""AgeDOB""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Birth Date For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Birth Date, from which their Age will be calculated and displayed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""City""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is City For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their City, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""State""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is State For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their State, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Country""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Country For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to choose their Country, to be viewed in their profile and in each Topic or Reply they post." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""aim""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is the AIM Option For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off features that allow users to enter their AIM username... then for other users to send them messages and/or add them to their buddy list." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""icq""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is the ICQ Option For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off features that allow users to enter their ICQ number... then for other users to send them ICQ messages and/or see if they are online." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""msn""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is the MSN Option For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off features that allow users to enter their MSN username... then for other users to view their MSN Username." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""yahoo""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is the YAHOO Option For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Turns On/Off features that allow users to enter their YAHOO username... then for other users to send them messages and/or add them to their buddy list." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Occupation""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Occupation For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Occupation, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Homepages""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Homepages For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to display their homepage link by their name on each post and in their Profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""FavLinks""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Favorite Links For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter 2 of their Favorite Links, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""MStatus""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Marital Status For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Marital Status, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Bio""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Bio For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Bio, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Hobbies""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Hobbies For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Hobbies, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""LNews""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Latest News For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Latest News, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""Quote""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Quote For?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Allow your users to enter their Quote, to be viewed in their profile." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "ranks"
		arrStarColors = ("Gold|Silver|Bronze|Orange|Red|Purple|Blue|Cyan|Green")
		arrIconStarColors = array(strIconStarGold,strIconStarSilver,strIconStarBronze,strIconStarOrange,strIconStarRed,strIconStarPurple,strIconStarBlue,strIconStarCyan,strIconStarGreen)
		strStarColor = split(arrStarColors, "|")

		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""ShowRank""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Showing Ranks?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                <ol>" & vbNewLine & _
				"                <li>Don't Show Any</li>" & vbNewLine & _
				"                <li>Show Rank Only</li>" & vbNewLine & _
				"                <li>Show Stars Only</li>" & vbNewLine & _
				"                <li>Show Both Stars and Rank</li>" & vbNewLine & _
				"                </ol>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""RankColor""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Color of Stars?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                You can change the color of stars that show up for each rank of member. (only when the Stars function is turned on)" & vbNewLine & _
				"                Available colors for the stars:<br /><br />" & vbNewLine
		for c = 0 to ubound(strStarColor)
			Response.Write	"                " & getCurrentIcon(arrIconStarColors(c),"","align=""absmiddle""") & "&nbsp;&nbsp;" & strStarColor(c)
			if c <> ubound(strStarColor) then Response.Write("<br />" & vbNewLine) else Response.Write(vbNewLine)
		next
		Response.Write	"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "datetime"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""timetype""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Time Display?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Choose 24Hr to display all times in military (24 hour) format or 12Hr to display all times in 12 hour format appended with an AM or PM depending on whether it's before or after midday. Default is 24 hour." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""TimeAdjust""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Time Adjustment?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Enter either a positive or negative integer value between +12 and 0 and -12. This may come in handy if you are located in one part of the world, and your server is in another, and you need the time displayed in the forum to be converted to a local time for you! (Default value is 0, meaning no adjustment)" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr> " & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""datetype""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Date Display?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Choose the format you wish all dates to be displayed in. Default is 12/31/2000 (US Short)." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "email"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""email""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does E-mail do?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Disabling the E-mail function will turn off any features that involve sending mail. If you don't have an SMTP server of any type, you will want to turn this feature off. If you do have an SMTP (mail) server, however, then also select the type of server you have from the dropdown menu." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""mailserver""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is a Mail Server Address?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                The mail server address is the actual domain name that resolves your mail server. This could be something like:<br />" & vbNewLine & _
				"                <b>mail.snitz.com</b><br />" & vbNewLine & _
				"                or it could be the same address as the web server:<br />" & vbNewLine & _
				"                <b>www.snitz.com</b><br />" & vbNewLine & _
				"                Either way, don't put the <b>http://</b> on it.<br />" & vbNewLine & _
				"                <br />" & vbNewLine & _
				"                <font color=""" & strHiLiteFontColor & """><b>NOTE:</b> If you are using CDONTS as a mail server type, you do not need to fill in this field.</font>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""sender""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Administrator E-mail Address?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This address is referenced by the forums in a couple ways.<br />" & vbNewLine & _
				"                <ol>" & vbNewLine & _
				"                <li>When mail is sent, it is sent from this user E-mail Account.</li>" & vbNewLine & _
				"                <li>This user is also the point of contact given if there is a problem with these forums.</li>" & vbNewLine & _
				"                </ol>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""UniqueEmail""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Unique E-mail Address?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Do you want to require each user to have their own E-mail Address?" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""EmailVal""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>E-mail Validation?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Do you want to require each user to validate their E-mail Address when they first Register and anytime they change their E-mail Address?<br /><br />The user will receive an E-mail with a link in it that will validate that the E-mail Address they entered is a valid E-mail Address." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""RestrictReg""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Restrict Registration?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This allows you to choose who is able to register on your forum by approving or rejecting their registration.<br /><br /><b>Note:</b> You must have the E-mail Validation option turned On to use this feature." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""LogonForMail""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Require Logon for sending Mail?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Do you require a user to be logged on before being able to use the <i>Send Topic To a Friend</i> or <i>E-mail Poster</i> options?" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
	case "colors"
		Response.Write	"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""fontfacetype""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>Font Face Type?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Font Face Type changes the way the text in your forum looks. You may want to change this option to match that of the rest of your web site. Some standards are:" & vbNewLine & _
				"                <ul>" & vbNewLine & _
				"                <li>Arial (nice, clean, legible font)</li>" & vbNewLine & _
				"                <li>Courier (a typewriter font)</li>" & vbNewLine & _
				"                <li>Helvetica (another clean, legible font)</li>" & vbNewLine & _
				"                <li>Sans Serif (Arial & Helvetica are variants of Sans Serif)</li>" & vbNewLine & _
				"                <li>Times New Roman (a book-type font)</li>" & vbNewLine & _
				"                <li>Verdana (another clean, legible font) (default)</li>" & vbNewLine & _
				"                </ul>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""fontsize""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What does Font Size mean?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                <ul>" & vbNewLine & _
				"                <li>None = Use Browser Default</li>" & vbNewLine & _
				"                <li>1 = 8 point font <b>X-Small</b> (default footer size)</li>" & vbNewLine & _
				"                <li>2 = 10 point font <b>Small</b> (default font size)</li>" & vbNewLine & _
				"                <li>3 = 12 point font <b>Normal</b></li>" & vbNewLine & _
				"                <li>4 = 14 point font <b>Large</b> (default header size)</li>" & vbNewLine & _
				"                <li>5 = 18 point font <b>X-Large</b></li>" & vbNewLine & _
				"                <li>6 = 24 point font <b>XX-Large</b></li>" & vbNewLine & _
				"                <li>7 = 36 point font <b>XXX-Large</b></li>" & vbNewLine & _
				"                </ul>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""colors""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What colors may I use?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """>" & vbNewLine & _
				"                <p><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                There are a lot of different colors you can choose from, all of which are listed below:</p>" & vbNewLine & _
				"                <blockquote><pre><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """>" & vbNewLine & _
				"                <font color=""aliceblue"">aliceblue</font>" & vbNewLine & _
				"                <font color=""antiquewhite"">antiquewhite</font>" & vbNewLine & _
				"                <font color=""aqua"">aqua</font>" & vbNewLine & _
				"                <font color=""aquamarine"">aquamarine</font>" & vbNewLine & _
				"                <font color=""azure"">azure</font>" & vbNewLine & _
				"                <font color=""beige"">beige</font>" & vbNewLine & _
				"                <font color=""bisque"">bisque</font>" & vbNewLine & _
				"                <font color=""black"">black</font>" & vbNewLine & _
				"                <font color=""blanchedalmond"">blanchedalmond</font>" & vbNewLine & _
				"                <font color=""blue"">blue</font>" & vbNewLine & _
				"                <font color=""blueviolet"">blueviolet</font>" & vbNewLine & _
				"                <font color=""brown"">brown</font>" & vbNewLine & _
				"                <font color=""burlywood"">burlywood</font>" & vbNewLine & _
				"                <font color=""cadetblue"">cadetblue</font>" & vbNewLine & _
				"                <font color=""chartreuse"">chartreuse</font>" & vbNewLine & _
				"                <font color=""chocolate"">chocolate</font>" & vbNewLine & _
				"                <font color=""coral"">coral</font>" & vbNewLine & _
				"                <font color=""cornflowerblue"">cornflowerblue</font>" & vbNewLine & _
				"                <font color=""cornsilk"">cornsilk</font>" & vbNewLine & _
				"                <font color=""cyan"">cyan</font>" & vbNewLine & _
				"                <font color=""darkblue"">darkblue</font>" & vbNewLine & _
				"                <font color=""darkcyan"">darkcyan</font>" & vbNewLine & _
				"                <font color=""darkgoldenrod"">darkgoldenrod</font>" & vbNewLine & _
				"                <font color=""darkgray"">darkgray</font>" & vbNewLine & _
				"                <font color=""darkgreen"">darkgreen</font>" & vbNewLine & _
				"                <font color=""darkkhaki"">darkkhaki</font>" & vbNewLine & _
				"                <font color=""darkmagenta"">darkmagenta</font>" & vbNewLine & _
				"                <font color=""darkolivegreen"">darkolivegreen</font>" & vbNewLine & _
				"                <font color=""darkorange"">darkorange</font>" & vbNewLine & _
				"                <font color=""darkorchid"">darkorchid</font>" & vbNewLine & _
				"                <font color=""darkred"">darkred</font>" & vbNewLine & _
				"                <font color=""darksalmon"">darksalmon</font>" & vbNewLine & _
				"                <font color=""darkseagreen"">darkseagreen</font>" & vbNewLine & _
				"                <font color=""darkslateblue"">darkslateblue</font>" & vbNewLine & _
				"                <font color=""darkslategray"">darkslategray</font>" & vbNewLine & _
				"                <font color=""darkturquoise"">darkturquoise</font>" & vbNewLine & _
				"                <font color=""darkviolet"">darkviolet</font>" & vbNewLine & _
				"                <font color=""deeppink"">deeppink</font>" & vbNewLine & _
				"                <font color=""deepskyblue"">deepskyblue</font>" & vbNewLine & _
				"                <font color=""dimgray"">dimgray</font>" & vbNewLine & _
				"                <font color=""dodgerblue"">dodgerblue</font>" & vbNewLine & _
				"                <font color=""firebrick"">firebrick</font>" & vbNewLine & _
				"                <font color=""floralwhite"">floralwhite</font>" & vbNewLine & _
				"                <font color=""forestgreen"">forestgreen</font>" & vbNewLine & _
				"                <font color=""gainsboro"">gainsboro</font>" & vbNewLine & _
				"                <font color=""ghostwhite"">ghostwhite</font>" & vbNewLine & _
				"                <font color=""gold"">gold</font>" & vbNewLine & _
				"                <font color=""goldenrod"">goldenrod</font>" & vbNewLine & _
				"                <font color=""gray"">gray</font>" & vbNewLine & _
				"                <font color=""green"">green</font>" & vbNewLine & _
				"                <font color=""greenyellow"">greenyellow</font>" & vbNewLine & _
				"                <font color=""honeydew"">honeydew</font>" & vbNewLine & _
				"                <font color=""hotpink"">hotpink</font>" & vbNewLine & _
				"                <font color=""indianred"">indianred</font>" & vbNewLine & _
				"                <font color=""ivory"">ivory</font>" & vbNewLine & _
				"                <font color=""khaki"">khaki</font>" & vbNewLine & _
				"                <font color=""lavender"">lavender</font>" & vbNewLine & _
				"                <font color=""lavenderblush"">lavenderblush</font>" & vbNewLine & _
				"                <font color=""lawngreen"">lawngreen</font>" & vbNewLine & _
				"                <font color=""lemonchiffon"">lemonchiffon</font>" & vbNewLine & _
				"                <font color=""lightblue"">lightblue</font>" & vbNewLine & _
				"                <font color=""lightcoral"">lightcoral</font>" & vbNewLine & _
				"                <font color=""lightcyan"">lightcyan</font>" & vbNewLine & _
				"                <font color=""lightgoldenrod"">lightgoldenrod</font>" & vbNewLine & _
				"                <font color=""lightgoldenrodyellow"">lightgoldenrodyellow</font>" & vbNewLine & _
				"                <font color=""lightgray"">lightgray</font>" & vbNewLine & _
				"                <font color=""lightgreen"">lightgreen</font>" & vbNewLine & _
				"                <font color=""lightpink"">lightpink</font>" & vbNewLine & _
				"                <font color=""lightsalmon"">lightsalmon</font>" & vbNewLine & _
				"                <font color=""lightseagreen"">lightseagreen</font>" & vbNewLine & _
				"                <font color=""lightskyblue"">lightskyblue</font>" & vbNewLine & _
				"                <font color=""lightslateblue"">lightslateblue</font>" & vbNewLine & _
				"                <font color=""lightslategray"">lightslategray</font>" & vbNewLine & _
				"                <font color=""lightsteelblue"">lightsteelblue</font>" & vbNewLine & _
				"                <font color=""lightyellow"">lightyellow</font>" & vbNewLine & _
				"                <font color=""limegreen"">limegreen</font>" & vbNewLine & _
				"                <font color=""linen"">linen</font>" & vbNewLine & _
				"                <font color=""magenta"">magenta</font>" & vbNewLine & _
				"                <font color=""maroon"">maroon</font>" & vbNewLine & _
				"                <font color=""mediumaquamarine"">mediumaquamarine</font>" & vbNewLine & _
				"                <font color=""mediumblue"">mediumblue</font>" & vbNewLine & _
				"                <font color=""mediumorchid"">mediumorchid</font>" & vbNewLine & _
				"                <font color=""mediumpurple"">mediumpurple</font>" & vbNewLine & _
				"                <font color=""mediumseagreen"">mediumseagreen</font>" & vbNewLine & _
				"                <font color=""mediumslateblue"">mediumslateblue</font>" & vbNewLine & _
				"                <font color=""mediumspringgreen"">mediumspringgreen</font>" & vbNewLine & _
				"                <font color=""mediumturquoise"">mediumturquoise</font>" & vbNewLine & _
				"                <font color=""mediumvioletred"">mediumvioletred</font>" & vbNewLine & _
				"                <font color=""midnightblue"">midnightblue</font>" & vbNewLine & _
				"                <font color=""mintcream"">mintcream</font>" & vbNewLine & _
				"                <font color=""mistyrose"">mistyrose</font>" & vbNewLine & _
				"                <font color=""moccasin"">moccasin</font>" & vbNewLine & _
				"                <font color=""navajowhite"">navajowhite</font>" & vbNewLine & _
				"                <font color=""navy"">navy</font>" & vbNewLine & _
				"                <font color=""navyblue"">navyblue</font>" & vbNewLine & _
				"                <font color=""oldlace"">oldlace</font>" & vbNewLine & _
				"                <font color=""olivedrab"">olivedrab</font>" & vbNewLine & _
				"                <font color=""orange"">orange</font>" & vbNewLine & _
				"                <font color=""orangered"">orangered</font>" & vbNewLine & _
				"                <font color=""orchid"">orchid</font>" & vbNewLine & _
				"                <font color=""palegoldenrod"">palegoldenrod</font>" & vbNewLine & _
				"                <font color=""palegreen"">palegreen</font>" & vbNewLine & _
				"                <font color=""paleturquoise"">paleturquoise</font>" & vbNewLine & _
				"                <font color=""palevioletred"">palevioletred</font>" & vbNewLine & _
				"                <font color=""papayawhip"">papayawhip</font>" & vbNewLine & _
				"                <font color=""peachpuff"">peachpuff</font>" & vbNewLine & _
				"                <font color=""peru"">peru</font>" & vbNewLine & _
				"                <font color=""pink"">pink</font>" & vbNewLine & _
				"                <font color=""plum"">plum</font>" & vbNewLine & _
				"                <font color=""powderblue"">powderblue</font>" & vbNewLine & _
				"                <font color=""purple"">purple</font>" & vbNewLine & _
				"                <font color=""red"">red</font>" & vbNewLine & _
				"                <font color=""rosybrown"">rosybrown</font>" & vbNewLine & _
				"                <font color=""royalblue"">royalblue</font>" & vbNewLine & _
				"                <font color=""saddlebrown"">saddlebrown</font>" & vbNewLine & _
				"                <font color=""salmon"">salmon</font>" & vbNewLine & _
				"                <font color=""sandybrown"">sandybrown</font>" & vbNewLine & _
				"                <font color=""seagreen"">seagreen</font>" & vbNewLine & _
				"                <font color=""seashell"">seashell</font>" & vbNewLine & _
				"                <font color=""sienna"">sienna</font>" & vbNewLine & _
				"                <font color=""skyblue"">skyblue</font>" & vbNewLine & _
				"                <font color=""slateblue"">slateblue</font>" & vbNewLine & _
				"                <font color=""slategray"">slategray</font>" & vbNewLine & _
				"                <font color=""snow"">snow</font>" & vbNewLine & _
				"                <font color=""springgreen"">springgreen</font>" & vbNewLine & _
				"                <font color=""steelblue"">steelblue</font>" & vbNewLine & _
				"                <font color=""tan"">tan</font>" & vbNewLine & _
				"                <font color=""thistle"">thistle</font>" & vbNewLine & _
				"                <font color=""tomato"">tomato</font>" & vbNewLine & _
				"                <font color=""turquoise"">turquoise</font>" & vbNewLine & _
				"                <font color=""violet"">violet</font>" & vbNewLine & _
				"                <font color=""violetred"">violetred</font>" & vbNewLine & _
				"                <font color=""wheat"">wheat</font>" & vbNewLine & _
				"                <font color=""white"">white</font>" & vbNewLine & _
				"                <font color=""whitesmoke"">whitesmoke</font>" & vbNewLine & _
				"                <font color=""yellow"">yellow</font>" & vbNewLine & _
				"                <font color=""yellowgreen"">yellowgreen</font>" & vbNewLine & _
				"                </font></pre></blockquote>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""fontdecorations""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What are Font Decorations?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                <ul>" & vbNewLine & _
				"                <li>none</li>" & vbNewLine & _
				"                <li>blink</li>" & vbNewLine & _
				"                <li>line-through</li>" & vbNewLine & _
				"                <li>overline</li>" & vbNewLine & _
				"                <li>underline</li>" & vbNewLine & _
				"                </ul>" & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""pagebgimage""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is Page Background Image URL?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                Enter the URL to the location of the background image you would like for your forum." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""columnwidth""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>How does Column Width Work?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                This sets the width of the column in question. It is not recommended that you change this unless you really know what your doing." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strCategoryCellColor & """><a name=""nowrap""></a><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """ ><b>What is NOWRAP?</b></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine & _
				"              <tr>" & vbNewLine & _
				"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strFooterFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
				"                NOWRAP prevents the text in a column from auto wrapping. This could be bad if you have people posting long strings of text in the right column (message box), reason being: this would cause an awful long horizontal scroll bar in most cases." & vbNewLine & _
				"                <a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></font></td>" & vbNewLine & _
				"              </tr>" & vbNewLine
end select
Response.Write	"            </table>" & vbNewLine & _
		"          </td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine
%>
<!--#INCLUDE FILE="inc_footer_short.asp" -->