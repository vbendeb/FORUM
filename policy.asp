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
		"          " & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Правила Регистрации и Пользования</font></td>" & vbNewLine & _
		"        </tr>" & vbNewLine & _
		"      </table>" & vbNewLine

if strProhibitNewMembers <> "1" then
	Response.Write	"      <table border=""0"" width=""100%"" cellspacing=""0"" cellpadding=""0"" align=""center"">" & vbNewLine & _
			"        <tr>" & vbNewLine & _
			"          <td bgcolor=""" & strTableBorderColor & """>" & vbNewLine & _
			"            <table border=""0"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strCategoryCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strCategoryFontColor & """><b>Процесс Регистрации и Правила Пользования " & strForumTitle & "</b></font></td>" & vbNewLine & _
			"              </tr>" & vbNewLine & _
			"              <tr>" & vbNewLine & _
			"                <td bgcolor=""" & strForumCellColor & """><font face=""" & strDefaultFontFace & """ size=""" & strDefaultFontSize & """ color=""" & strForumFontColor & """>" & vbNewLine & _
			"                <p><U>Процесс Регистрации</U><BR> " & _
			"<font color=""" & strHiLiteFontColor & """>Внимание: Вам не удастся активировать свой логин и получить доступ к форуму без работающего E-mail адреса. Если Вы не получили регистрационного сообщения, повторите регистрацию сначала, на этот раз с правильным адресом E-mail.  Возможно Вам придется выбрать другое имя логина при повторной регистрации.</p></font>" & vbNewLine & _
			"1. После заполнения анкеты и нажатия кнопки &quot;Регистрация&quot; на следующей странице Ваш запрос посылается на MOCT.<BR> " & _
			"2. Вебмастер подтвердит Вашу регистрацию на форуме (это может занять несколько дней!).<BR> " & _
			"3. В результате подтверждения на Ваш E-mail адрес будет выслано регистрационное сообщение с специальной ссылкой на Форум.<BR> " & _
			"4. Нажатие (клик) на ссылку в регистрационном сообщении активизирует Вашу регистрацию.<BR> " & _
			"5. Шаг №4 подтверждает правильность Вашего E-mail адреса. Если адрес неправильный, процесс остановится на №2.<BR> " & _
			" <BR> " & _
			"                <p><U>Проблемы При Регистрации</U><BR> " & _
			"Самая частая проблема - &quot;незаконченная регистрация&quot;. Это может произойти в следующих случаях:<BR> " & _
			"1. Указание неправильного E-Mail адреса (очепятки).<BR> " & _
			"2. Cистемное сообщение посланное на указанный адрес не было получено адресатом (недоставка почты).<BR> " & _
			"3. Cистемное сообщение посланное на указанный адрес не было подтверждено адресатом.<BR> " & _
			"В любом из этих случаев, новый НИК &quot;зависнет&quot; и повторение регистрации под этим же ником будет запрещено. В этом случае, пошлите E-Mail на адрес <A href='mailto:administrator@moct.org'>administrator@moct.org</A> с обязательным указанием НИКа и E-Mail адреса.<BR> " & _
			" <BR> " & _
			"                <p><U>Правила Пользования</U><BR> " & _
			"Администрация Форума не может постоянно контролировать содержание сообщений и поэтому не в состоянии оградить " & _
			"Вас от чего-то что может Вам не понравиться или даже оскорбить. Тем не менее, мы будем удалять всё что мы посчитаем неприемлемым и вредным для нашего Форума. " & _
			"Нажатием на кнопку &quot;Регистрация&quot; Вы выражаете свое согласие с Правилами " & _
			"и принимаете на себя полную ответственность за Ваши действия на Форуме и за информацию Вы публикуете. " & _
			"Вы также обещаете НЕ публиковать никаких незаконных метериалов и НЕ нарушать законов авторского права. " & _
			"Вы также обязуетесь НЕ публиковать никаких сообщений вульгарного, оскорбительного, ненавистнического или сексуального характера. </p>" & vbNewLine & _
			"                <p><U>Персональная Информация</U><BR> " & _
			"<font color=""" & strHiLiteFontColor & """>Персональная информация содержащаяся в Вашем персональном профиле НЕОБХОДИМА И ОБЯЗАТЕЛЬНА для всех участников МОСТа и Форума.  Наша главная задача - помочь выпускникам ГНИ и их друзьям найти друг друга, что невозможно в случае сохранения ими анонимности.  Ваша персональная информация не может и не будет использована ни в каких других целях!  В целях защиты, анонимные пользователи будут удаляться из списка участников и лишаться доступа к Форуму.  На случай злоупотребления Вашим или нашим доверием, наша система считывает и хранит первый и последний IP адрес каждого пользователя, по которому можно легко выйти на провайдера, институт, школу или организацию которой этот адрес принадлежит. Участники могут добавить или изменить информацию о себе используя свой логин и пароль для изменения данных в их персональном Профиле. В случае утери пароля, Вы можете послать запрос на E-Mail адрес " & _
			"  <span class=""spnMessageText""><a href=""mailto:" & strSender & """>" & strSender & "</a></span>.</font>" & _
			"</ol>" & _
			"</p>" & vbNewLine & _
			"                <p>Если Вы согласны с Правилами Форума изложенными выше, пожалуйста нажмите кнопку &quot;Перейти к регистрации&quot;, в противном случае нажмите кнопку &quot;Отмена&quot;.</p>" & vbNewLine & _
			"                <hr size=""1"">" & vbNewLine & _
			"                  <table align=""center"" border=""0"">" & vbNewLine & _
			"                    <tbody>" & vbNewLine & _
			"                      <tr>" & vbNewLine & _
			"                        <td>" & vbNewLine & _
			"                        <form action=""register.asp?mode=Register"" id=""form1"" method=""post"" name=""form1"">" & vbNewLine & _
			"                        <input name=""Refer"" type=""hidden"" value=""" & Request.ServerVariables("HTTP_REFERER") & """>" & vbNewLine & _
			"                        <input name=""Submit"" type=""Submit"" value=""Перейти к регистрации"">" & vbNewLine & _
			"                        </form>" & vbNewLine & _
			"                        </td>" & vbNewLine & _
			"                        <td>" & vbNewLine & _
			"                        <form action=""JavaScript:history.go(-1)"" id=""form2"" method=""post"" name=""form2"">" & vbNewLine & _
			"                        <input name=""Submit"" type=""Submit"" value=""Отмена"">" & vbNewLine & _
			"                        </form>" & vbNewLine & _
			"                        </td>" & vbNewLine & _
			"                      </tr>" & vbNewLine & _
			"                    </tbody>" & vbNewLine & _
			"                  </table>" & vbNewLine & _
			"                <hr size=""1"">" & vbNewLine & _
			"                <p>По всем вопросам Форума " & _
			"Вы можете связаться с нами по адресу: " & _
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