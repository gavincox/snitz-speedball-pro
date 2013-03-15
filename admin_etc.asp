<%
'##################################################################################################
'## Snitz Forums 2000 v3.4.07
'##################################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##		   Huw Reddick and Richard Kinser
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
'##################################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Forum Cleanup Tools<br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
strToDo          = Request.QueryString("action")
strTempMemID     = cLng(Request.QueryString("member_id"))
strTempForumID   = cLng(Request.QueryString("forum_id"))
strTempDelMember = Request.QueryString("delmember")
strTempDelReply  = Request.QueryString("delreply")
Select Case LCase(strToDo)
	Case "deletememtopics"
		if Request.QueryString("c") = "t" then
			DelMemTopicExec
		else
			response.write "<br><center>Go ahead with requested operation?<br><a href=""admin_etc.asp?c=t&delreply=" & strTempDelReply & "&action=" & strToDo & "&member_id=" & strTempMemID & "&forum_id=" & strTempForumID & "&delmember=" & strTempDelMember & """>Yes</a> | <a href=""admin_etc.asp"">No</a></center><br><br>"
		end if
	Case "delforumtopics"
		if Request.QueryString("c") = "t" then
			DelForumTopicsExec
		else
			response.write "<br><center>Go ahead with requested operation?<br><a href=""admin_etc.asp?c=t&action=" & strToDo & "&delreply=" & strTempDelReply & "&member_id=" & strTempMemID & "&forum_id=" & strTempForumID & "&delmember=" & strTempDelMember & """>Yes</a> | <a href=""admin_etc.asp"">No</a></center><br><br>"
		end if
	Case "deletememtopicsdone"
		Response.write "<br><center>All Done!<br>" & strLE & _
			"<a href=""admin_etc.asp"">Back to Forum Cleanup Tools</a><br></center><br>"
	Case else
		response.write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Delete Topics by Member</b></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE
		DelMemTopicHTML()
		Response.write "</tr>" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Delete Topics by Forum</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE
		DelForumTopicsHTML()
		response.write "</tr>" & strLE &_
			"</table>"
End Select
Call WriteFooter
Sub DelForumTopicsHTML()
	response.write "<td class=""r"">" & strLE & _
		"<form method=""get"" action=""admin_etc.asp"">" & strLE & _
		"<input type=""hidden"" name=""action"" value=""delforumtopics"">" & strLE & _
		"Forum&nbsp;</td><td><select name=""forum_id"">" & strLE
	strsql = "SELECT FORUM_ID, F_SUBJECT FROM " & strTablePrefix & "FORUM"
	set fRS = my_Conn.execute(strsql)
	if not(fRS.eof) then
		do while not(fRS.eof)
			response.write "<option value=""" & fRS("FORUM_ID") & """>" & fRS("F_SUBJECT") & "</option>" & strLE
			fRS.eof ' movenext
		loop
	else
		response.write "<option value=""0"">No Forums Found</option>" & strLE
	end if
	set fRS = nothing
	response.write "</select></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Delete Topics""></td>" & strLE & _
		"</tr>" & strLE & _
		"</form>" & strLE & _
		"</td>"
End Sub
Sub DelMemTopicHTML()
	response.write "<td class=""r""><form method=""get"" action=""admin_etc.asp"">" & strLE & _
		"<input type=""hidden"" name=""action"" value=""deletememtopics""> Member&nbsp;</td>" & strLE & _
		"<td><select name=""member_id"">" & strLE
'## Delete Topics by Member
	strsql = "SELECT MEMBER_ID, M_NAME, M_LEVEL FROM " & strMemberTablePrefix & "MEMBERS ORDER BY M_NAME ASC"
	set aRS = my_Conn.execute(strsql)
	if not(aRS.eof) then
		do while not aRS.eof
			if not(aRS("M_LEVEL") > 2) then response.write "<option value=""" & aRS("MEMBER_ID") & """>" & aRS("M_NAME") & "</option>" & strLE
			aRS.movenext
		loop
	else
		response.write "<option value=""0"">No Members Found</option>"
	end if
	set aRS = nothing
	response.write "</select></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""vat r"">Delete&nbsp;</td>" & strLE & _
		"<td><input type=""checkbox"" name=""delreply"" value=""delreply""> Member Replies<br>" & strLE & _
		"<input type=""checkbox"" name=""delmember"" value=""delmember""> Member Profile</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Delete Topics""></td>" & strLE & _
		"</tr>" & strLE & _
		"</form></td>"
End Sub
Sub DelMemTopicExec()
	if strTempDelReply = "delreply" and not(strTempMemID = 1) then
		strsql = "SELECT TOPIC_ID, REPLY_ID, FORUM_ID, R_STATUS FROM " & strTablePrefix & "REPLY WHERE R_AUTHOR=" & strTempMemID
		set rs = my_Conn.execute(strsql)
		do while not (rs.eof)
			strTempRSreplyID = rs("REPLY_ID")
			strTempRStopicID = rs("TOPIC_ID")
			Reply_Status     = rs("R_STATUS")
			Forum_ID         = rs("FORUM_ID")
			strsql           = "DELETE FROM " & strTablePrefix & "REPLY WHERE REPLY_ID=" & strTempRSreplyID
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
			strSql = "SELECT REPLY_ID, R_DATE, R_AUTHOR, R_STATUS"
			strSql = strSql & " FROM " & strTablePrefix & "REPLY "
			strSql = strSql & " WHERE TOPIC_ID = " & strTempRStopicID & " "
			strSql = strSql & " AND R_STATUS <= 1 "
			strSql = strSql & " ORDER BY R_DATE DESC"
			set trs = my_Conn.Execute (strSql)
			if not(trs.eof or trs.bof) then
				strLast_Post_Reply_ID = trs("REPLY_ID")
				strLast_Post = trs("R_DATE")
				strLast_Post_Author = trs("R_AUTHOR")
			end if
			if (trs.eof or trs.bof) or IsNull(strLast_Post) or IsNull(strLast_Post_Author) then
				set rs2 = Server.CreateObject("ADODB.Recordset")
				strSql = "SELECT T_AUTHOR, T_DATE "
				strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
				strSql = strSql & " WHERE TOPIC_ID = " & strTempRStopicID & " "
				set rs2 = my_Conn.Execute (strSql)
				strLast_Post_Reply_ID = 0
				strLast_Post = rs2("T_DATE")
				strLast_Post_Author = rs2("T_AUTHOR")
				rs2.Close
				set rs2 = nothing
			end if
				if Reply_Status <= 1 then
					strSql = "UPDATE " & strTablePrefix & "TOPICS "
					strSql = strSql & " SET T_REPLIES = T_REPLIES - 1 "
					if strLast_Post <> "" then
						strSql = strSql & ", T_LAST_POST = '" & strLast_Post & "'"
						if strLast_Post_Author <> "" then
							strSql = strSql & ", T_LAST_POST_AUTHOR = " & strLast_Post_Author & ""
						end if
					end if
					strSql = strSql & ", T_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID & ""
					strSql = strSql & " WHERE TOPIC_ID = " & strTempRStopicID
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
					'## Forum_SQL - Get last_post and last_post_author for Forum
					strSql = "SELECT TOPIC_ID, T_LAST_POST, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID "
					strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
					strSql = strSql & " WHERE FORUM_ID = " & Forum_ID & " "
					strSql = strSql & " ORDER BY T_LAST_POST DESC"
					set srs = my_Conn.Execute (strSql)
					if not srs.eof then
						strLast_Post = srs("T_LAST_POST")
						strLast_Post_Author = srs("T_LAST_POST_AUTHOR")
						strLast_Post_Topic_ID = srs("TOPIC_ID")
						strLast_Post_Reply_ID = srs("T_LAST_POST_REPLY_ID")
					else
						strLast_Post = ""
						strLast_Post_Author = "NULL"
						strLast_Post_Topic_ID = 0
						strLast_Post_Reply_ID = 0
					end if
					srs.Close
					set srs = nothing
						strSql =  "UPDATE " & strTablePrefix & "FORUM "
						strSql = strSql & " SET F_COUNT = F_COUNT - 1 "
						strSql = strSql & ", F_LAST_POST = '" & strLast_Post & "'"
						strSql = strSql & ", F_LAST_POST_AUTHOR = " & strLast_Post_Author
						strSql = strSql & ", F_LAST_POST_TOPIC_ID = " & strLast_Post_Topic_ID
						strSql = strSql & ", F_LAST_POST_REPLY_ID = " & strLast_Post_Reply_ID
						strSql = strSql & " WHERE FORUM_ID = " & Forum_ID
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
						'## FORUM_SQL - Decrease count of total replies in Totals table by 1
						strSql = "UPDATE " & strTablePrefix & "TOTALS "
						strSql = strSql & " SET P_COUNT = P_COUNT - 1 "
						my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				else
					strSql = "UPDATE " & strTablePrefix & "TOPICS "
					strSql = strSql & " SET T_UREPLIES = T_UREPLIES - 1 "
					strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID
					my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				end if
			rs.movenext
		loop
		set rs = nothing
	end if
	'## Delete all stuff associated with topics being deleted...
	strsql = "SELECT TOPIC_ID FROM " & strTablePrefix & "TOPICS WHERE T_AUTHOR=" & strTempMemID
	set rs = my_conn.execute(strsql)
	if not(rs.eof) then
		do while not(rs.eof)
			strsql = "DELETE FROM " & strTablePrefix & "REPLY WHERE TOPIC_ID=" & rs("TOPIC_ID")
			my_conn.execute(strsql)
			strsql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS WHERE TOPIC_ID=" & rs("TOPIC_ID")
			my_conn.execute(strsql)
			rs.movenext
		loop
	end if
	set rs = nothing
	'## End delete stuff
	strsql = "DELETE FROM " & strTablePrefix & "TOPICS WHERE T_AUTHOR=" & strTempMemID
	my_Conn.execute(strsql)
	if strTempDelMember = "delmember" and not(strTempMemID = 1) then
		strsql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS WHERE MEMBER_ID = " & strTempMemID
		my_conn.execute(strsql)
	end if
	Response.write("<br><center>The Requested Operation is in progress, please wait...</center>" & strLE & _
		"<meta http-equiv=""Refresh"" content=""1; URL=admin_count.asp?comeback=etc""><br><br>")
End Sub
Sub DelForumTopicsExec()
	'Delete all the replies for the indicated forum
	strsql = "DELETE FROM " & strTablePrefix & "REPLY WHERE FORUM_ID=" & strTempForumID
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	'Delete all the topics from the forum
	strsql = "DELETE FROM " & strTablePrefix & "TOPICS WHERE FORUM_ID=" & strTempForumID
	my_conn.execute(strsql)
	'Update the last post info in the FORUM table
	strsql = "UPDATE " & strTablePrefix & "FORUM "
	strsql = strsql & "SET F_LAST_POST='', F_LAST_POST_AUTHOR=0, F_LAST_POST_REPLY_ID=0, F_LAST_POST_TOPIC_ID=0 "
	strsql = strsql & "WHERE FORUM_ID=" & strTempForumID
	my_conn.execute(strsql)
	Response.write("<br><center>The Requested Operation is in progress, please wait...</center>" & strLE & _
		"<meta http-equiv=""Refresh"" content=""1; URL=admin_count.asp?comeback=etc""><br><br>")
End Sub
%>
