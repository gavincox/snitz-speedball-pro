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
If IsNumeric(Request.QueryString("Forum")) Then Forum_ID   = cLng(Request.QueryString("Forum")) Else Forum_ID = 0
If IsNumeric(Request.QueryString("userid")) Then User_ID   = cLng(Request.QueryString("userid")) Else User_ID = 0
If IsNumeric(Request.QueryString("action")) Then Action_ID = cInt(Request.QueryString("action")) Else Action_ID = 0
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp"-->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
if Forum_ID = 0 then
	txtMessage = "Select a forum to edit moderators for that forum"
else
	if User_ID = 0 then
		txtMessage = "Select a user to grant/revoke moderator powers for that user.	Users in bold are currently moderators of this forum."
	else
		if Action_ID = 0 then
			txtMessage = "Select an action for this user"
		else
			txtMessage = "Action Successful"
		end if
	end if
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Moderator Configuration<br><br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE & _
	"<table class=""admin"">" & strLE & _
	"<tr>" & strLE & _
	"<th><b>Moderator Configuration</b></th>" & strLE & _
	"</tr>" & strLE
	if txtMessage <> "" Then Response.Write "<tr><th>" & txtMessage & "</th></tr>" & strLE
if Forum_ID = 0 then
	'## Forum_SQL
	strSql = "SELECT C.CAT_ORDER, C.CAT_NAME, F.CAT_ID, F.FORUM_ID, F.F_ORDER, F.F_SUBJECT " &_
	" FROM " & strTablePrefix & "CATEGORY C, " & strTablePrefix & "FORUM F" &_
	" WHERE C.CAT_ID = F.CAT_ID "
	strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT ASC;"
	set rs = my_Conn.Execute(strSql)
	if rs.eof or rs.bof then
		'nothing
	else
		iOldCat = 0
		do until rs.EOF
			iNewCat = rs("CAT_ID")
			if iNewCat <> iOldCat Then
				Response.Write "<tr><th><b>" & rs("CAT_NAME") & "</b></th></tr>" & strLE
				iOldCat = iNewCat
			end if
			Response.Write "<tr><td>" & getCurrentIcon(strIconFolder,"","class=""vam""") & "&nbsp;<span class=""smt"">" & strLE & _
				"<a href=""admin_moderators.asp?forum=" & rs("FORUM_ID") & """>" & rs("F_SUBJECT") & "</a>" & strLE & _
				"</span></td></tr>" & strLE
			rs.MoveNext
		loop
	end if
else
	if Action_ID = 0 then
		if User_ID = 0 then
			'## Forum_SQL
			strSql = "SELECT " & strMemberTablePrefix & "MEMBERS.MEMBER_ID, " & strMemberTablePrefix & "MEMBERS.M_NAME "
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_LEVEL > 1 "
			strSql = strSql & " AND   " & strMemberTablePrefix & "MEMBERS.M_STATUS = " & 1
			strSql = strSql & " ORDER BY " & strMemberTablePrefix & "MEMBERS.M_NAME ASC;"
			set rs = my_Conn.Execute(strSql)
			Response.Write "<tr>" & strLE & _
			"<td>" & strLE & _
				"<ul>" & strLE
			do until rs.EOF
				Response.Write "<li>"
				if chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("<b>")
				Response.Write "<span class=""smt""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & rs("MEMBER_ID")& """>" & rs("M_NAME") & "</a></span>"
				If chkForumModerator(Forum_ID, rs("M_NAME")) then Response.Write("</b>")
				Response.Write "</li>" & strLE
				rs.MoveNext
			loop
			Response.Write "</ul>" & strLE
		else
			'## Forum_SQL
			strSql = "SELECT " & strTablePrefix & "MODERATOR.FORUM_ID, " & strTablePrefix & "MODERATOR.MEMBER_ID, " & strTablePrefix & "MODERATOR.MOD_TYPE "
			strSql = strSql & " FROM " & strTablePrefix & "MODERATOR "
			strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID & " "
			strSql = strSql & " AND " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "
			set rs = my_Conn.Execute(strSql)
			if rs.EOF then
				Response.Write "<tr>" & strLE & _
					"<td class=""c"">" & strLE & _
					"The selected user is NOT currently a moderator of the selected forum<br>" & strLE & _
					"<br>" & strLE & _
					"<span class=""smt""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=1"">Assign moderator permission</a></span>" & strLE & _
					"<br>" & strLE
			else
				Response.Write "<tr>" & strLE & _
					"<td class=""c"">" & strLE & _
					"The selected user IS currently a moderator of the selected forum<br>" & strLE & _
					"<br>" & strLE & _
					"<span class=""smt""><a href=""admin_moderators.asp?forum=" & Forum_ID & "&UserID=" & User_ID & "&action=2"">Remove moderator permission</a></span>" & strLE & _
					"<br>" & strLE
			end if
		end if
	else
		select case Action_ID
			case 1
				'## Forum_SQL
				strSql = "INSERT INTO " & strTablePrefix & "MODERATOR "
				strSql = strSql & "(FORUM_ID"
				strSql = strSql & ", MEMBER_ID"
				strSql = strSql & ") VALUES ("
				strSql = strSql & Forum_ID
				strSql = strSql & ", " & User_ID
				strSql = strSql & ")"
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				Response.Write "<tr>" & strLE & _
					"<td class=""c"">" & strLE & _
					"The selected user is now a moderator of the selected forum<br>" & strLE & _
					"<br>" & strLE & _
					"<span class=""smt""><a href=""admin_moderators.asp"">Back to Moderator Options</a></span>" & strLE & _
					"<br>" & strLE
			case 2
				'## Forum_SQL
				strSql = "DELETE FROM " & strTablePrefix & "MODERATOR "
				strSql = strSql & " WHERE " & strTablePrefix & "MODERATOR.FORUM_ID = " & Forum_ID & " "
				strSql = strSql & " AND   " & strTablePrefix & "MODERATOR.MEMBER_ID = " & User_ID
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				Response.Write "<tr>" & strLE & _
					"<td class=""c"">" & strLE & _
					"The selected user's moderator status in the selected forum has been removed<br>" & strLE & _
					"<br>" & strLE & _
					"<span class=""smt""><a href=""admin_moderators.asp"">Back to Moderator Options</a></span>" & strLE & _
					"<br>" & strLE
		end select
	end if
end if
Response.Write "</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
Call WriteFooter
Response.End
%>
