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
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & server.urlencode(scriptname(ubound(scriptname)) & "?" & request.querystring)
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Forum Deletion/Archival<br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
strWhatToDo = request.querystring("action")
if strWhatToDo = "" then strWhatToDo = "default"
Select Case strWhatToDo
	Case "default"
		Response.Write "<table class=""admin"">" & strLE & _
			"<th><b>Administrative Forum Archive Functions</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<th><b>Forum Options</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr class=""vat"">" & strLE & _
			"<td><ul>" & strLE & _
			"<li class=""smt""><a href=""admin_forums.asp?action=archive"">Archive topics from a forum</a></li>" & strLE & _
			"<li class=""smt""><a href=""admin_forums.asp?action=deletearchive"">Delete selected topics from an archive</a></li>" & strLE & _
			"<li class=""smt""><a href=""admin_forums_schedule.asp"">Configure Archive Reminder</a></li>" & strLE & _
			"<li class=""smt""><a href=""admin_forums.asp?action=delete"">Delete <b>all</b> topics from a forum</a></li>" & strLE
		if strDBType = "access" and Instr(19,strConnString,"Jet",1) > 0 then Response.write("<li class=""smt""><a href=""admin_compactdb.asp"">Compact Database</a></li>" & strLE)
		Response.Write "</ul></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><a href=""admin_home.asp"">Back to Admin Home</a></p>" & strLE
	Case "delete" ' ################## DELETE
		Response.Write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Administrative Forum Delete Functions</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Delete Topics</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE
		strForumIDN = request.querystring("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" then
			strsql = "SELECT CAT_ID, FORUM_ID, F_L_DELETE, F_SUBJECT,F_DELETE_SCHED FROM " & strTablePrefix & "FORUM ORDER BY CAT_ID, F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)
			thisCat = 0
			if drs.eof then
				Response.Write "<td colspan=""2"">No Forums Found!</td>" & strLE & _
					"</tr>" & strLE
			else
				Response.Write "<td colspan=""2""><ul>" & strLE & _
					"<li class=""smt""><a href=""admin_forums.asp?action=delete&id=-1"">All Forums</a></li>" & strLE & _
					"<li><a href=""javascript:document.delTopic.submit()"">Selected Forums</a></li>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE & _
					"<form name=""delTopic"" action=""admin_forums.asp"">" & strLE & _
					"<input type=""hidden"" value=""delete"" name=""action"">" & strLE
				do until drs.eof
					lastDeleted = drs("F_L_DELETE")
					schedDays = drs("F_DELETE_SCHED")
					if (IsNull(lastDeleted)) or (lastDeleted = "") then
						delete_date = "N/A"
						overdue = 0
					else
						needDelete = (DateAdd("d",schedDays+7,strToDate(lastDeleted)))
						if (strForumTimeAdjust > needDelete) and (schedDays > 0) then
							overdue = true
							delete_date = "<span class=""hlfc"">Deletion Overdue</span>"
						else
							overdue = false
							delete_date = StrToDate(lastDeleted)
						end if
					end if
					if thisCat <> drs("CAT_ID") then response.write "<tr><td colspan=""2"">&nbsp;</td></tr>"
					Response.Write "<tr>" & strLE & _
						"<td><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """"
					if overdue then Response.Write(" checked")
					Response.Write ">&nbsp;<span class=""smt""><a href=""admin_forums.asp?action=delete&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a></span></td>" & strLE & _
						"<td class=""r"">Last delete date: " & delete_date & "</td>" & strLE & _
						"</tr>" & strLE
					thisCat = drs("CAT_ID")
					drs.movenext
				loop
				Response.Write "</form>" & strLE
			end if
			set drs = nothing
			Response.Write "</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE
		elseif request.querystring("confirm") = "true" then
			Response.Write "<center><span class=""dff dfs ffc"">All Topics in selected Forum/s have been Deleted.</span></center><br>" & strLE
			Call subdeletestuff(strForumIDN)
		elseif request.querystring("confirm") = "" then
			Response.Write "<center><span class=""dff dfs ffc"">Are you sure you want to delete <b>ALL</b> topics"
			if Request.QueryString("id") = "-1" then Response.Write(" in <b>ALL</b> forums? ") else Response.Write(" in the selected forums? ")
			Response.Write "This is <B><STRONG>NOT</STRONG></B> reversible.<br><br>" & strLE & _
				"<span class=""smt""><a href=""admin_forums.asp?action=delete&id=" & strForumIDN & "&confirm=true"">Yes</a></span> | <span class=""smt""><a href=""admin_forums.asp?action=delete&id=" & strForumIDN & "&confirm=false"">No</a></span></span></center><br>" & strLE
		elseif request.querystring("confirm") = "false" then
			Response.Write "<center><span class=""dff dfs ffc"">Topics in selected Forum/s have NOT been deleted.</span></center><br><br>" & strLE
		end if
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<br>" & strLE & _
			"<center><span class=""dff dfs dfc""><a href=""admin_forums.asp"">Back to Forums Administration</a></span></center><br>" & strLE & _
			"<br>" & strLE
	Case "archive" '################ ARCHIVE
		Response.Write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Administrative Forum Archive Functions</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Archive all topics</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE
		strForumIDN = request("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" then
			strsql  = "Select CAT_ID, FORUM_ID, F_L_ARCHIVE, F_SUBJECT,F_ARCHIVE_SCHED from " & strTablePrefix & "FORUM WHERE F_TYPE = 0 ORDER BY CAT_ID, F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)
			thisCat = 0
			if drs.eof then
				Response.Write "<td colspan=""2"">No Forums Found</td>" & strLE & _
					"</tr>" & strLE
			else
				Response.Write "<td colspan=""2""><ul><li class=""smt""><a href=""admin_forums.asp?action=archive&id=-1"">All Forums</a></li>" & strLE & _
					"<li class=""smt""><a href=""javascript:document.arcTopic.submit()"">Selected Forums</a></li></ul></td>" & strLE & _
					"</tr>" & strLE & _
					"<form name=""arcTopic"" action=""admin_forums.asp"">" & strLE & _
					"<input type=""hidden"" value=""archive"" name=""action"">" & strLE
				do until drs.eof
			           	lastArchived = drs("F_L_ARCHIVE")
					schedDays = drs("F_ARCHIVE_SCHED")
					if (IsNull(lastArchived)) or (lastArchived = "") then
						archive_date = "Not archived"
						overdue = 0
					else
						needArchive = (DateAdd("d",schedDays+7,strToDate(lastArchived)))
						if (strForumTimeAdjust > needArchive) and (schedDays > 0) then
							overdue = true
							archive_date = "<span class=""hlfc"">Archiving Overdue</span>"
						else
							overdue = false
							archive_date = StrToDate(lastArchived)
						end if
					end if
					if thisCat <> drs("CAT_ID") then response.write "<tr><td colspan=""2"">&nbsp;</td></tr>" & strLE
					Response.Write "<tr>" & strLE & _
						"<td class=""smt""><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """"
					if overdue then Response.Write(" checked")
					Response.Write """>&nbsp;<a href=""admin_forums.asp?action=archive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a></td>" & strLE & _
						"<td class=""r"">Last archive date: " & archive_date & "</td>" & strLE & _
						"</tr>" & strLE
					thisCat = drs("Cat_ID")
					drs.movenext
				loop
				Response.Write "</form>" & strLE
			end if
			set drs = nothing
			Response.Write "</table>" & strLE
		elseif strForumIDN <> "" then
			if request.querystring("confirm") = "" then
				Response.Write "<form method=""post"" action=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=no"">" & strLE & _
						"<br>" & strLE & _
						"<span class=""dff dfs ffc"">Archive Topics which are older than:</span>&nbsp;&nbsp;" & strLE & _
						"<select name=""archiveolderthan"" size=""1"">" & strLE
				for counter = 1 to 6
					Response.Write "                    	<option value=""" & DateToStr(DateAdd("m", -counter, now())) & """>" & counter & " Month"
					if counter > 1 then response.write("s")
					Response.Write "</option>" & strLE
				next
				Response.Write "                      	<option value=""" & DateToStr(DateAdd("m", -12, now())) & """>One Year</option>" & strLE & _
						"</select>" & strLE & _
						"                      &nbsp;&nbsp;" & strLE & _
						"<input type=""submit"" value=""Archive"">" & strLE & _
						"</form>" & strLE
			elseif request.querystring("confirm") = "no" then
				Response.Write "<center><span class=""dff dfs ffc"">Are you sure you want to archive these topics?<br><br>" & strLE & _
						"<span class=""smt""><a href=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=yes&date=" & request.form("archiveolderthan") & """>Yes</a></span> | <span class=""smt""><a href=""admin_forums.asp?action=archive&id=" & strForumIDN & "&confirm=cancel"">No</a></span></span></center><br>" & strLE
			elseif request.querystring("confirm") = "yes" then
				If chkDateFormat(request.querystring("date")) Then Call subarchivestuff(request.querystring("date"))
			elseif request.querystring("confirm") = "cancel" then
				Response.Write "<span class=""dff dfs ffc"">Archiving Cancelled.</span><br><br>" & strLE
			end if
			Response.Write "<br>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE
		end if
		Response.Write "<br>" & strLE & _
				"<center><span class=""dff dfs dfc""><a href=""admin_forums.asp"">Back to Forums Administration</a></span></center><br>" & strLE & _
				"<br>" & strLE
	Case "deletearchive" '######################## DELETE ARCHIVED
		Response.Write "<table width=""75%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
				"<tr>" & strLE & _
				"<td>" & strLE & _
				"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""ccc""><span class=""dff dfs cfc""><b>Administrative Forum Archive Functions</b></span></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""fcc vat""><span class=""dff dfs ffc""><b>Delete archived topics:</b></span></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""fcc vat c"">" & strLE
		strForumIDN = request.querystring("id")
		strForumIDN = Server.URLEncode(strForumIDN)
		if strForumIDN = "" and request.querystring("confirm") = "" then
			Response.Write "<table width=""100%"">" & strLE & _
					"<tr>" & strLE & _
					"<td colspan=""2""><span class=""dff dfs ffc"">Select a forum from which to delete archived topics</span><br></td>" & strLE & _
					"</tr>" & strLE
   			strSql = "SELECT " & strTablePrefix & "FORUM.CAT_ID, "
		    	strSql = strSql & strTablePrefix & "FORUM.FORUM_ID, "
		    	strSql = strSql & strTablePrefix & "FORUM.F_L_DELETE, "
		    	strSql = strSql & strTablePrefix & "FORUM.F_DELETE_SCHED, "
		    	strSql = strSql & strTablePrefix & "FORUM.F_SUBJECT "
		    	strSql = strSql & " FROM " & strTablePrefix & "FORUM, " & strArchiveTablePrefix & "TOPICS "
		    	strSql = strSql & " WHERE " & strTablePrefix & "FORUM.FORUM_ID = " & strArchiveTablePrefix & "TOPICS.FORUM_ID "
		    	strSql = strSql & " ORDER BY " & strTablePrefix & "FORUM.CAT_ID DESC, " & strTablePrefix & "FORUM.F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)
			thisCat = 0
			thisForum = 0
			if drs.eof then
				Response.Write "<tr>" & strLE & _
						"<td colspan=""2""><span class=""dff dfs ffc""><b>No Forums Found!</b></span></td>" & strLE & _
						"</tr>" & strLE
		        else
				Response.Write "<tr>" & strLE & _
						"<td colspan=""2""><span class=""dff dfs ffc""><li><span class=""smt""><a href=""admin_forums.asp?action=deletearchive&id=-1"">All Forums</a></span></span></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td colspan=""2""><span class=""dff dfs ffc""><li><span class=""smt""><a href=""javascript:document.delTopic.submit()"">Selected Forums</a></span></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td colspan=""2""><span class=""dff dfs ffc"">&nbsp;</td>" & strLE & _
						"</tr>" & strLE & _
						"<form name=""delTopic"" action=""admin_forums.asp"">" & strLE & _
						"<input type=""hidden"" value=""deletearchive"" name= ""action"">" & strLE
				do until drs.eof
					if thisForum <> drs("FORUM_ID") then
						thisForum = drs("FORUM_ID")
				           	lastDeleted = drs("F_L_DELETE")
						schedDays = drs("F_DELETE_SCHED")
						if (IsNull(lastDeleted)) or (lastDeleted = "") then
							delete_date = "N/A"
							overdue = 0
						else
							needDelete = (DateAdd("d",schedDays+7,strToDate(lastDeleted)))
							if (strForumTimeAdjust > needDelete) and (schedDays > 0) then
								overdue = true
								delete_date = "<span class=""hlfc"">Deletion Overdue</span>"
							else
								overdue = false
								delete_date = StrToDate(lastDeleted)
							end if
						end if
						if thisCat <> drs("CAT_ID") then
							response.write "<tr><td colspan=""2"">&nbsp;</td></tr>"
							thisCat = drs("CAT_ID")
						end if
						Response.Write "<tr>" & strLE & _
								"<td><span class=""dff dfs ffc""><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & ""
						if overdue then Response.Write(" checked")
						Response.Write """><span class=""smt""><a href=""admin_forums.asp?action=deletearchive&id=" & drs("FORUM_ID") & """>" & drs("F_SUBJECT") & "</a></span></span></td>" & strLE & _
								"<td><span class=""dff dfs ffc""> Last delete date: " & delete_date & "</span></td>" & strLE & _
								"</tr>" & strLE
					end if
					drs.movenext
				loop
				Response.Write "</form>" & strLE
			end if
			set drs = nothing
				Response.Write "</table>" & strLE
		elseif request.querystring("id") <> "" and request.querystring("confirm") = "" then
			Response.Write 	"<center><span class=""dff dfs ffc"">Select how many months old the Topics should be that you wish to delete</span></center>" & strLE & _
				"<form method=""post"" action=""admin_forums.asp?action=deletearchive&id=" & strForumIDN & "&confirm=no"">" & strLE & _
				"<center><span class=""dff dfs ffc"">Delete archived Topics which are older than:</span><br>" & strLE & _
				"<select name=""archiveolderthan"" size=""1"">" & strLE
			for counter = 1 to 6
				Response.Write " <option value=""" & DateToStr(DateAdd("m", -counter, now())) & """>" & counter & " Month"
				if counter > 1 then Response.Write("s")
				Response.Write "</option>" & strLE
			next
			Response.Write "<option value=""" & DateToStr(DateAdd("m", -12, now())) & """>One Year</option>" & strLE & _
				"</select>" & strLE & _
				"&nbsp;&nbsp;" & strLE & _
				"<input type=""submit"" value=""Delete""></center>" & strLE & _
				"</form>" & strLE
     		elseif request.querystring("id") <> "" and request.querystring("confirm") = "no" then
     			Response.Write "<center><span class=""dff dfs ffc"">Are you sure you want to delete these topics from the archive?<br><br>" & strLE & _
					"<span class=""smt""><a href=""admin_forums.asp?action=deletearchive&id=" & strForumIDN & "&confirm=yes&date=" & request.form("archiveolderthan") & """>Yes</a></span> | <span class=""smt""><a href=""admin_forums.asp?action=delete&confirm=false&id=" & strForumIDN & """>No</a></span></span></center><br>" & strLE
     		elseif strForumIDN <> "" and request.querystring("confirm") = "yes" then
	            	Response.Write "<center><span class=""dff dfs ffc"">Topics older than " & StrToDate(request.querystring("date")) & " have been deleted from the selected archive forum.</span></center><br>" & strLE
     			If chkDateFormat(request.querystring("date")) Then Call subdeletearchivetopics(strForumIDN, request.querystring("date"))
		end if
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<br>" & strLE & _
			"<center><span class=""dff dfs dfc""><a href=""admin_forums.asp"">Back to Forums Administration</a></span></center><br>" & strLE & _
			"<br>" & strLE
end Select
Sub subDeleteArchiveTopics(strForum_id, strDateOlderThan)
	Dim fIDSQL
	'#### create FORUM_ID clause
	rqID = request("id")
	'rqID = strForum_id
        on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then
			fIDSQL = " AND FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " AND FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	strsql = "DELETE FROM " & strArchiveTablePrefix & "TOPICS WHERE T_LAST_POST < '" & strDateOlderThan & "'" & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	strsql = "DELETE FROM " & strArchiveTablePrefix & "REPLY WHERE R_DATE < '" & strDateOlderThan & "'" & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	Call subdoupdates()
End Sub
Sub subArchiveStuff(fdateolderthan)
	set Server2 = Server
	Server2.ScriptTimeout = 10000
	Dim fIDSQL
	Dim drs,delRep
	Set drs = CreateObject("ADODB.Recordset")
	Set delRep = CreateObject("ADODB.Recordset")
	Set drs.ActiveConnection = my_conn
	'#### create FORUM_ID clause
	rqID = request("id")
    	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then
			fIDSQL = " AND " & strTablePrefix & "TOPICS.FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " AND " & strTablePrefix & "TOPICS.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	'#### Get the replies to Archive
	strSql = "SELECT T_DATE, " & strTablePrefix & "REPLY.* FROM " & strTablePrefix & "REPLY LEFT OUTER JOIN " & strTablePrefix & "TOPICS " &_
		 "ON " & strTablePrefix & "REPLY.TOPIC_ID = " & strTablePrefix & "TOPICS.TOPIC_ID " &_
		 " WHERE T_LAST_POST < '" & fdateolderthan & "'" & fIDSQL
	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "
	drs.Open strsql, my_conn, adOpenStatic, adLockOptimistic, adCmdText
	'#### Archive the Replies
	if drs.eof then
    		response.write("<center><span class=""dff dfs ffc"">No Replies were Archived: none found</span></center><br>" & strLE)
	else
        	i = 0
		response.write("<span class=""dff ffs ffc"">")
		do until drs.eof
			if isnull(drs("R_LAST_EDITBY")) then
				intR_LAST_EDITBY = "NULL"
			else
				intR_LAST_EDITBY = drs("R_LAST_EDITBY")
			end if
        		strsqlvalues = "" & drs("CAT_ID") & ", " & drs("FORUM_ID") & ", " & drs("TOPIC_ID") & ", " & drs("REPLY_ID")
		        strsqlvalues = strsqlvalues & ", " & drs("R_AUTHOR") & ", '" & chkstring(drs("R_MESSAGE"),"archive")
	       	        strsqlvalues = strsqlvalues & "', '" & drs("R_DATE") & "', '" & drs("R_IP") & "'"  & ", " & drs("R_STATUS")
			strSqlvalues = strsqlvalues & ", '" & drs("R_LAST_EDIT") & "', " & intR_LAST_EDITBY & ", " & drs("R_SIG") & " "
	                strsql = "INSERT INTO " & strArchiveTablePrefix & "REPLY (CAT_ID, FORUM_ID, TOPIC_ID, REPLY_ID, R_AUTHOR, R_MESSAGE, R_DATE, R_IP, R_STATUS, R_LAST_EDIT, R_LAST_EDITBY, R_SIG)"
		        strsql = strsql & " VALUES (" & strsqlvalues & ")"
			response.write(".")
			'Response.Write(strSql)
			'Response.End
			my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	           	drs.movenext
			i = i + 1
			if i = 100 then
				response.write("<br>")
				i = 0
			end if
			'#### Delete Original
		Loop
		response.write("</span>" & strLE)
		drs.movefirst
		do while not drs.eof
			strsql = "select * from " & strTablePrefix & "REPLY WHERE REPLY_ID = " & drs("REPLY_ID")
			delrep.Open strsql, my_conn, adOpenStatic, adLockOptimistic, adCmdText
			delrep.delete
			delrep.close
			drs.movenext
		loop
		response.write("<center><span class=""dff dfs ffc"">All replies to Topics older than " & strToDate(fdateolderthan) & " were archived</span></center><br>" & strLE)
	end if
	'#### Update FORUM archive date
	strsql = "UPDATE " & strTablePrefix & "FORUM SET F_L_ARCHIVE= '" & fdateolderthan & "'"
	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then
			strSQL = strSql & " WHERE FORUM_ID=" & rqID
		end if
		err.clear
	else
		strSQL = strSql & " WHERE FORUM_ID IN (" & rqID & ")"
		err.clear
	end if
	on error goto 0
'	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	'#### Get the TOPICS to Archive
	strsql = "SELECT CAT_ID,FORUM_ID,TOPIC_ID,T_SUBJECT,T_AUTHOR,T_REPLIES,T_UREPLIES,T_VIEW_COUNT,T_LAST_POST,T_DATE,T_LAST_POSTER,T_IP,T_LAST_POST_AUTHOR,T_LAST_POST_REPLY_ID,T_LAST_EDIT,T_LAST_EDITBY,T_STICKY,T_SIG,T_MESSAGE FROM " & strTablePrefix & "TOPICS WHERE T_LAST_POST < '" & fdateolderthan & "'" & fIDSQL
	strSQL = strSQL & " AND T_ARCHIVE_FLAG <> 0 "
	set drs = my_conn.execute(strsql)
	'#### Archive the Topics
   	if drs.eof then
       		response.write("<center><span class=""dff dfs ffc"">No Topics were Archived: none found</span></center><br>" & strLE)
	else
	       	i = 0
       		do until drs.eof
       			strSQL = "SELECT TOPIC_ID FROM " & strArchiveTablePrefix & "TOPICS WHERE TOPIC_ID=" & drs("TOPIC_ID")
			set rsTcheck = my_conn.execute(strSQL)
			if isnull(drs("T_LAST_EDITBY")) then
				intT_LAST_EDITBY = "NULL"
			else
				intT_LAST_EDITBY = drs("T_LAST_EDITBY")
			end if
			if isnull(drs("T_LAST_POST_REPLY_ID")) then
				intT_LAST_POST_REPLY_ID = "NULL"
			else
				intT_LAST_POST_REPLY_ID = drs("T_LAST_POST_REPLY_ID")
			end if
			if isnull(drs("T_UREPLIES")) then
				intT_UREPLIES = "NULL"
				intT_UREPLIEScnt = 0
			else
				intT_UREPLIES = drs("T_UREPLIES")
				intT_UREPLIEScnt = drs("T_UREPLIES")
			end if
			if rsTcheck.eof then
				err.clear
				strsqlvalues = "" & drs("CAT_ID") & ", " & drs("FORUM_ID") & ", " & drs("TOPIC_ID") & ", " & 0
		           	strsqlvalues = strsqlvalues & ", '" & chkstring(drs("T_SUBJECT"),"archive") & "', '" & chkstring(drs("T_MESSAGE"),"archive")
		           	strsqlvalues = strsqlvalues & "', " & drs("T_AUTHOR") & ", " & drs("T_REPLIES") & ", " & intT_UREPLIES & ", " & drs("T_VIEW_COUNT")
	        	   	strsqlvalues = strsqlvalues & ", '" & drs("T_LAST_POST") & "', '" & drs("T_DATE") & "', " & drs("T_LAST_POSTER")
	           		strsqlvalues = strsqlvalues & ", '" & drs("T_IP") & "', " & drs("T_LAST_POST_AUTHOR") & ", " & intT_LAST_POST_REPLY_ID & ", '" & drs("T_LAST_EDIT")
				strsqlvalues = strsqlvalues & "', " & intT_LAST_EDITBY & ", " & drs("T_STICKY") & ", " & drs("T_SIG") & " "
		       		strsql = "INSERT INTO " & strArchiveTablePrefix & "TOPICS (CAT_ID, FORUM_ID, TOPIC_ID, T_STATUS, T_SUBJECT, T_MESSAGE, T_AUTHOR, T_REPLIES, T_UREPLIES, T_VIEW_COUNT, T_LAST_POST, T_DATE, T_LAST_POSTER, T_IP, T_LAST_POST_AUTHOR, T_LAST_POST_REPLY_ID, T_LAST_EDIT, T_LAST_EDITBY, T_STICKY, T_SIG)"
				strsql = strsql & " VALUES (" & strsqlvalues & ")"
				'Response.Write strSql
				'Response.End
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
				msg = "<center><span class=""dff dfs ffc"">All topics older than " & strToDate(fdateolderthan) & " were archived</span></center><br>" & strLE
			else
		       		strsql = "UPDATE " & strArchiveTablePrefix & "TOPICS SET " &_
					"T_STATUS = " & 0 &_
					", T_SUBJECT = '" & chkstring(drs("T_SUBJECT"),"archive") & "'" &_
					", T_MESSAGE = '" & chkstring(drs("T_MESSAGE"),"archive") & "'" &_
					", T_REPLIES = T_REPLIES + " & drs("T_REPLIES") &_
					", T_UREPLIES = T_UREPLIES + " & intT_UREPLIEScnt &_
					", T_VIEW_COUNT = T_VIEW_COUNT + " & drs("T_VIEW_COUNT") &_
					", T_LAST_POST = '" & drs("T_LAST_POST") & "'" &_
					", T_LAST_POST_AUTHOR = " & drs("T_LAST_POST_AUTHOR") &_
					", T_LAST_POST_REPLY_ID = " & intT_LAST_POST_REPLY_ID & _
					", T_LAST_EDIT = '" & drs("T_LAST_EDIT") & "'" & _
					", T_LAST_EDITBY = " & intT_LAST_EDITBY & _
					", T_STICKY = " & drs("T_STICKY") & _
					", T_SIG = " & drs("T_SIG") & _
					" WHERE TOPIC_ID = " & drs("TOPIC_ID")
 	            		response.write("<span class=""dff ffs ffc"">." & strLE)
				my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
				msg = "<br><center>Topic exists, Stats Updated......</center></span>" & strLE
			end if
		        Response.Write msg
			'#### Delete originals
			if i > 100 then
				i = 0
				response.write("<br>" & strLE)
			end if
			i = i + 1
			'## Forum_SQL - Delete any subscriptions to this topic
			strSql = "DELETE FROM " & strTablePrefix & "SUBSCRIPTIONS "
			strSql = strSql & " WHERE TOPIC_ID = " & drs("TOPIC_ID")
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
           drs.movenext
	Loop
	drs.close
	strSql = "DELETE FROM " & strTablePrefix & "TOPICS WHERE T_LAST_POST < '" & fdateolderthan & "' " & fIDSQL
	strSqL = strSqL & " AND T_ARCHIVE_FLAG <> 0 "
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
    End if
    Call subdoupdates()
    'response.write("<br><center><a href=""admin_forums.asp"">Click Here</a> to return to Forums Delete/Archive Admin</center><br>" & strLE)
End Sub
Sub subdeletestuff(fstrid)
	Dim fIDSQL
'#### create FORUM_ID clause
	rqID = request("id")
    	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then
			fIDSQL = " WHERE FORUM_ID=" & rqID
		else
			fIDSQL = ""
		end if
		err.clear
	else
		fIDSQL = " WHERE FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	strsql = "DELETE FROM " & strTablePrefix & "TOPICS " & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	strsql = "DELETE FROM " & strTablePrefix & "REPLY " & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	'#### Update FORUM last delete posts date
	strsql = "UPDATE " & strTablePrefix & "FORUM SET F_L_DELETE= '" & DateToStr(now()) & "'"
	strsql = strsql & fIDSQL
	my_conn.execute(strsql),,adCmdText + adExecuteNoRecords
	Call subdoupdates()
End Sub
Sub subdoupdates()
	'#### create FORUM_ID clause
	rqID = request("id")
    	on error resume next
	testID = cLng(rqID)
	if err.number = 0 then
		if rqID <> "-1" then
			fIDSQL = " AND " & strTablePrefix & "FORUM.FORUM_ID=" & rqID
			fIDSQL2 = " WHERE " & strTablePrefix & "TOPICS.FORUM_ID=" & rqID
		else
			fIDSQL = ""
			fIDSQL2 = ""
		end if
		err.clear
	else
		fIDSQL = " AND " & strTablePrefix & "FORUM.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		fIDSQL2 = " WHERE " & strTablePrefix & "TOPICS.FORUM_ID IN (" & ChkString(rqID, "SQLString") & ")"
		err.clear
	end if
	on error goto 0
	Response.Write "<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""c"" colspan=""2""><p><b><span class=""dff ffs ffc"">Updating Counts</span></b><br></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat r""><span class=""dff ffs ffc"">Topics:</span></td>" & strLE & _
			"<td class=""vat""><span class=""dff ffs ffc"">"
	set rs = Server.CreateObject("ADODB.Recordset")
	set rs1 = Server.CreateObject("ADODB.Recordset")
	'## Forum_SQL - Get contents of the Forum table related to counting
	strSql = "SELECT FORUM_ID, F_TOPICS FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 " & fIDSQL
	rs.Open strSql, my_Conn
	if not(rs.EOF or rs.BOF) then
		rs.MoveFirst
		i = 0
		do until rs.EOF
			i = i + 1
			'## Forum_SQL - count total number of topics in each forum in Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
			set rs1 = my_Conn.Execute( strSql)
			if rs1.EOF or rs1.BOF then
				intF_TOPICS = 0
			else
				intF_TOPICS = rs1("cnt")
			end if
			rs1.Close
			'## Forum_SQL - count total number of archived topics in each forum in A_Topics table
			strSql = "SELECT count(FORUM_ID) AS cnt "
			strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
			set rs1 = my_Conn.Execute( strSql)
			if rs1.EOF or rs1.BOF then
				intF_A_TOPICS = 0
			else
				intF_A_TOPICS = rs1("cnt")
			end if
			rs1.Close
			strSql = "UPDATE " & strTablePrefix & "FORUM "
			strSql = strSql & " SET F_TOPICS = " & intF_TOPICS
			strSql = strSql & " , F_A_TOPICS = " & intF_A_TOPICS
			strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
			my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
			rs.MoveNext
			Response.Write "."
			if i = 80 then
				Response.Write "<br>"
				i = 0
			end if
		loop
	end if
	rs.Close
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat r""><span class=""dff ffs ffc"">Topic Replies:</span></td>" & strLE & _
			"<td class=""vat""><span class=""dff ffs ffc"">"
	'## Forum_SQL
	strSql = "SELECT TOPIC_ID, T_REPLIES FROM " & strTablePrefix & "TOPICS" & fIDSQL2
	rs.Open strSql, my_Conn
	i = 0
	do until rs.EOF
		i = i + 1
		'## Forum_SQL - count total number of replies in Topics table
		strSql = "SELECT count(REPLY_ID) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "REPLY "
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")
		rs1.Open strSql, my_Conn
		if rs1.EOF or rs1.BOF or (rs1("cnt") = 0) then
			intT_REPLIES = 0
		else
			intT_REPLIES = rs1("cnt")
		end if
		strSql = "UPDATE " & strTablePrefix & "TOPICS "
		strSql = strSql & " SET T_REPLIES = " & intT_REPLIES
		strSql = strSql & " WHERE TOPIC_ID = " & rs("TOPIC_ID")
		my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
		rs1.Close
		rs.MoveNext
		Response.Write "."
		if i = 80 then
			Response.Write "<br>"
			i = 0
		end if
	loop
	rs.Close
	Response.Write 	"</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat r""><span class=""dff ffs ffc"">Forum Replies:</span></td>" & strLE & _
			"<td class=""vat""><span class=""dff ffs ffc"">"
	'## Forum_SQL - Get values from Forum table needed to count replies
	strSql = "SELECT FORUM_ID, F_COUNT FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	rs.Open strSql, my_Conn, adOpenDynamic, adLockOptimistic, adCmdText
	do until rs.EOF
		'## Forum_SQL - Count total number of Replies
		strSql = "SELECT Sum(" & strTablePrefix & "TOPICS.T_REPLIES) AS SumOfT_REPLIES, Count(" & strTablePrefix & "TOPICS.T_REPLIES) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "TOPICS "
		strSql = strSql & " WHERE " & strTablePrefix & "TOPICS.FORUM_ID = " & rs("FORUM_ID")
		rs1.Open strSql, my_Conn
		if rs1.EOF or rs1.BOF then
			intF_COUNT = 0
			intF_TOPICS = 0
		else
			if IsNull(rs1("SumOfT_REPLIES")) then
				intF_COUNT = rs1("cnt")
			else
				intF_COUNT = CLng(rs1("cnt")) + CLng(rs1("SumOfT_REPLIES"))
			end if
			intF_TOPICS = rs1("cnt")
		end if
		if IsNull(intF_COUNT) then intF_COUNT = 0
		if IsNull(intF_TOPICS) then intF_TOPICS = 0
		rs1.Close
		'## Forum_SQL - Count total number of Archived Replies
		strSql = "SELECT Sum(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS SumOfT_A_REPLIES, Count(" & strTablePrefix & "A_TOPICS.T_REPLIES) AS cnt "
		strSql = strSql & " FROM " & strTablePrefix & "A_TOPICS "
		strSql = strSql & " WHERE " & strTablePrefix & "A_TOPICS.FORUM_ID = " & rs("FORUM_ID")
		rs1.Open strSql, my_Conn
		if rs1.EOF or rs1.BOF then
			intF_A_COUNT = 0
			intF_A_TOPICS = 0
		else
			if IsNull(rs1("SumOfT_A_REPLIES")) then
				intF_A_COUNT = rs1("cnt")
			else
				intF_A_COUNT = CLng(rs1("cnt")) + CLng(rs1("SumOfT_A_REPLIES"))
			end if
			intF_A_TOPICS = rs1("cnt")
		end if
		if IsNull(intF_A_COUNT) then intF_A_COUNT = 0
		if IsNull(intF_A_TOPICS) then intF_A_TOPICS = 0
		rs1.Close
		strSql = "UPDATE " & strTablePrefix & "FORUM "
		strSql = strSql & " SET F_COUNT = " & intF_COUNT
		strSql = strSql & ",  F_TOPICS = " & intF_TOPICS
		strSql = strSql & ",  F_A_COUNT = " & intF_A_COUNT
		strSql = strSql & ",  F_A_TOPICS = " & intF_A_TOPICS
		strSql = strSql & " WHERE FORUM_ID = " & rs("FORUM_ID")
		my_conn.execute(strSql),,adCmdText + adExecuteNoRecords
		rs.MoveNext
		Response.Write "."
		if i = 80 then
			Response.Write "<br>" & strLE
			i = 0
		end if
	loop
	rs.Close
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat r""><span class=""dff ffs ffc"">Totals:</span></td>" & strLE & _
			"<td class=""vat""><span class=""dff ffs ffc"">"
	'## Forum_SQL - Total of Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_TOPICS) "
	strSql = strSql & " AS SumOfF_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	rs.Open strSql, my_Conn
	if IsNull(RS("SumOfF_TOPICS")) then
		Response.Write "Total Topics: 0<br>" & strLE
		strSumOfF_TOPICS = 0
	else
		Response.Write "Total Topics: " & RS("SumOfF_TOPICS") & "<br>" & strLE
		strSumOfF_TOPICS = rs("SumOfF_TOPICS")
	end if
	rs.Close
	'## Forum_SQL - Total of Archived Topics
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_A_TOPICS) "
	strSql = strSql & " AS SumOfF_A_TOPICS "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	rs.Open strSql, my_Conn
	if IsNull(RS("SumOfF_A_TOPICS")) then
		Response.Write "Total Archived Topics: 0<br>" & strLE
		strSumOfF_A_TOPICS = 0
	else
		Response.Write "Total Archived Topics: " & RS("SumOfF_A_TOPICS") & "<br>" & strLE
		strSumOfF_A_TOPICS = rs("SumOfF_A_TOPICS")
	end if
	rs.Close
	'## Forum_SQL - Total all the replies for each topic
	strSql = "SELECT Sum(" & strTablePrefix & "FORUM.F_COUNT) "
	strSql = strSql & " AS SumOfF_COUNT "
	strSql = strSql & ", Sum(" & strTablePrefix & "FORUM.F_A_COUNT) "
	strSql = strSql & " AS SumOfF_A_COUNT "
	strSql = strSql & " FROM " & strTablePrefix & "FORUM WHERE F_TYPE <> 1 "
	set rs = my_Conn.Execute (strSql)
	if rs("SumOfF_COUNT") <> "" then
		Response.Write "Total Posts: " & RS("SumOfF_COUNT") & "<br>" & strLE
		strSumOfF_COUNT = rs("SumOfF_COUNT")
	else
		Response.Write "Total Posts: 0<br>" & strLE
		strSumOfF_COUNT = "0"
	end if
	if rs("SumOfF_A_COUNT") <> "" then
		Response.Write "Total Archived Posts: " & RS("SumOfF_A_COUNT") & "<br>" & strLE
		strSumOfF_A_COUNT = rs("SumOfF_A_COUNT")
	else
		Response.Write "Total Archived Posts: 0<br>" & strLE
		strSumOfF_A_COUNT = "0"
	end if
	set rs = nothing
	'## Forum_SQL - Write totals to the Totals table
	strSql = "UPDATE " & strTablePrefix & "TOTALS "
	strSql = strSql & " SET T_COUNT = " & strSumOfF_TOPICS
	strSql = strSql & ", P_COUNT = " & strSumOfF_COUNT
	strSql = strSql & ", T_A_COUNT = " & strSumOfF_A_TOPICS
	strSql = strSql & ", P_A_COUNT = " & strSumOfF_A_COUNT
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	Response.Write "</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2"">&nbsp;<br><b><span class=""dff ffs ffc"">Count Update Complete</span></b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>"
	set rs = nothing
	set rs1 = nothing
End Sub
Call WriteFooter
Response.End
%>
