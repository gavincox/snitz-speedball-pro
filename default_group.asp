<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##                       Huw Reddick and Richard Kinser
'##
'## This program is free software; you can redistribute it and/or
'## modify it under the terms of the GNU General Public License
'## as published by the Free Software Foundation; either version 2
'## of the License, or (at your option) any later version.
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
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
if strGroupCategories <> "1" then Response.Redirect("default.asp")
if strAutoLogon = 1 then
	if (ChkAccountReg() <> "1") then Response.Redirect "register.asp?mode=DoIt"
end if
strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_DESCRIPTION, GROUP_ICON"
strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
strSql = strSql & " ORDER BY GROUP_NAME ASC "
set rs = my_Conn.Execute (strSql)
Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""hcc vat nw c""><b><span class=""dff dfs hfc"">&nbsp;</span></b></td>" & strLE & _
	"<td class=""hcc vat nw c""><b><span class=""dff dfs hfc"">Discussion&nbsp;Groups</span></b></td>" & strLE & _
	"<td class=""hcc vat nw c""><b><span class=""dff dfs hfc"">Categories</span></b></td>" & strLE & _
	"<td class=""hcc vat nw c""><b><span class=""dff dfs hfc"">Forums</span></b></td>" & strLE & _
	"<td class=""hcc vat nw c""><b><span class=""dff dfs hfc"">Topics</span></b></td>" & strLE & _
	"</tr>" & strLE
if rs.EOF or rs.BOF then
	Response.Write "<tr>" & strLE & _
		"<td class=""ccc"" colspan="""
	if (strShowModerators = "1") or (mlev > 0 ) then Response.Write "6" else Response.Write "5"
	Response.Write """><span class=""dff cfc dfs vat""><b>No Categories/Forums Found</b></span></td>" & strLE
'	if (mlev = 4 or mlev = 3) then
	Response.Write "<td class=""ccc""><span class=""dff cfc dfs vat"">&nbsp;</span></td>" & strLE
'	end if
	Response.Write "</tr>" & strLE
else
	'rs.moveFirst
	do until rs.EOF
		if rs("GROUP_ID") = 1 then
			'do nothing
		else
			numCats   = 0
			numTopics = 0
			numPosts  = 0
			' how many categories ?
			strSql = "SELECT GROUP_ID, GROUP_CATID "
			strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
			strSql = strSql & " WHERE GROUP_ID = " & rs("GROUP_ID")
			strSql = strSql & " ORDER BY GROUP_ID ASC "
			set rsGroupCats = my_Conn.execute (strSql)
			if not rsGroupCats.eof then
				strSQLForum = "SELECT Count(CAT_ID) FROM " & strTablePrefix & "FORUM WHERE "
				strSQLTopic = "SELECT Count(CAT_ID) FROM " &  strTablePrefix & "TOPICS WHERE "
				first = 0
				do until rsGroupCats.eof
					numCats = numCats + 1
					if first = 0 then
						strSQLForum = strSQLForum & " CAT_ID =" & rsGroupCats("GROUP_CATID")
						strSQLTopic = strSQLTopic & " CAT_ID =" & rsGroupCats("GROUP_CATID")
						first = 1
					else
						strSQLForum = strSQLForum & " OR CAT_ID =" & rsGroupCats("GROUP_CATID")
						strSQLTopic = strSQLTopic & " OR CAT_ID =" & rsGroupCats("GROUP_CATID")
					end if
					rsGroupCats.MoveNext
				loop
				rsGroupCats.close
				set rsGroupCats = nothing
				set rsPostCount = my_Conn.execute (strSQLTopic)
				Select Case rsPostCount.eof
					Case False : NumTopics = rsPostCount(0)
					Case True  : NumTopics = 0
				End Select
				set rsPostCount = nothing
				set rsGroupForums = my_Conn.execute (strSqlForum)
				Select Case rsGroupForums.eof
					Case False : NumForums = rsGroupForums(0)
					Case True  : NumForums = 0
				End Select
				set rsGroupForums = nothing
			else
				NumCats   = 0
				NumForums = 0
				NumTopics = 0
			end if
			Response.Write "<tr>" & strLE & _
				"<td class=""fcc vat nw c"">"
			if instr(rs("GROUP_ICON"),".") then
				Response.Write getCurrentIcon(rs("GROUP_ICON") & "|20|20","","class=""vam""") & "</td>" & strLE
			else
				Response.Write getCurrentIcon(strIconGroupCategories,"","class=""vam""") & "</td>" & strLE
			end if
			Response.Write "<td class=""fcc vat l""><span class=""dff dfs ffc""><span class=""smt""><a href=""default.asp?group=" & cLng(rs("GROUP_ID")) & """>" & chkString(rs("GROUP_NAME"),"display") & "</a></span>"
			if rs("GROUP_DESCRIPTION") <> "" then Response.Write "<br><span class=""ffs"">" & formatStr(rs("GROUP_DESCRIPTION")) & "</span>"
			Response.Write "</span></td>" & strLE & _
				"<td class=""fcc vat nw c""><span class=""dff dfs ffc"">" & NumCats & "</span></td>" & strLE & _
				"<td class=""fcc vat nw c""><span class=""dff dfs ffc"">" & NumForums & "</span></td>" & strLE & _
				"<td class=""fcc vat nw c""><span class=""dff dfs ffc"">" & NumTopics & "</span></td>" & strLE & _
				"</tr>" & strLE
		end if
		rs.movenext
	loop
end if
rs.close
set rs = nothing
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table><br>" & strLE
Call WriteFooter
%>
