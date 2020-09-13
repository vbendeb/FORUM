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


'##############################################
'##                Do Counts                 ##
'##############################################

sub doPCount()
	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET P_COUNT = P_COUNT + 1"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub


sub doTCount()
	'## Forum_SQL - Updates the totals Table
	strSql ="UPDATE " & strTablePrefix & "TOTALS SET T_COUNT = T_COUNT + 1"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub


sub doUCount(sUser_Name)
	'## Forum_SQL - Update Total Post for user
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_POSTS = M_POSTS + 1 "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(sUser_Name, "SQLString") & "'"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub


sub doULastPost(sUser_Name)
	'## Forum_SQL - Updates the M_LASTPOSTDATE in the FORUM_MEMBERS table
	strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " SET M_LASTPOSTDATE = '" & DateToStr(strForumTimeAdjust) & "' "
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & ChkString(sUser_Name, "SQLString") & "'"

	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
end sub

%>