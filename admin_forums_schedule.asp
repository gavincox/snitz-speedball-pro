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
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Forum Deletion/Archival</td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE
strWhatToDo = request("action")
if strWhatToDo = "" then strWhatToDo = "default"
Select Case strWhatToDo
	Case "updateArchive"
		if Request("id") = "" or IsNull(Request("id")) then
			Response.Write "<p class=""c hfs"">There has been a problem!</p>" & strLE & _
				"<p class=""c hfs"">No Forums Selected!</p>" & strLE & _
				"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go back to correct the problem</a></p>" & strLE
			Call WriteFooter
			Response.End
		end if
		Response.Write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th><b>Administrative Forum Archive Schedule</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr class=""vat"">" & strLE & _
			"<td><ul>" & strLE
		reqID = split(Request.Form("id"), ",")
		for i = 0 to ubound(reqID)
			tmpStr = "archSched" & trim(reqID(i))
			if tmpStr = "" then tmpStr = NULL
			strSQL = "UPDATE " & strTablePrefix & "FORUM SET F_ARCHIVE_SCHED = " & cLng("0" & Request.Form(tmpStr))
			strSQL = strSQL & " WHERE FORUM_ID = " & cLng("0" & trim(reqID(i)))
			my_conn.execute(strSQL),,adCmdText + adExecuteNoRecords
			Response.Write "<li>Archive Schedule for <b>" & GetForumName(reqID(i)) & "</b> updated to " & Request.Form(tmpStr) & " days</li>" & strLE
		next
		Response.Write "</ul></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><a href=""admin_forums.asp"">Back to Forums Administration</a></p>" & strLE
	Case "default" '################ ARCHIVE
		Response.Write "<form name=""arcTopic"" action=""admin_forums_schedule.asp"" method=""post"">" & strLE & _
			"<input type=""hidden"" name=""action"" value=""updateArchive"">" & strLE & _
			"<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Administrative Forum Archive Functions</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Archive Reminder</b></th>" & strLE & _
			"</tr>" & strLE
		strForumIDN = request("id")
		if strForumIDN = "" then
			strsql = "SELECT CAT_ID, FORUM_ID, F_L_ARCHIVE,F_ARCHIVE_SCHED, F_SUBJECT FROM " & strTablePrefix & "FORUM ORDER BY CAT_ID, F_SUBJECT DESC"
			set drs = my_conn.execute(strsql)
			thisCat = 0
			if drs.eof then
				Response.Write "<tr><td colspan=""2"">No Forums Found!</td></tr>" & strLE
			else
				do until drs.eof
					if (IsNull(drs("F_L_ARCHIVE"))) or (drs("F_L_ARCHIVE") = "") then archive_date = "Not archived" else archive_date = StrToDate(drs("F_L_ARCHIVE"))
					Response.Write "<tr>" & strLE & _
						"<td><input type=""checkbox"" name=""id"" value=""" & drs("FORUM_ID") & """> " & drs("F_SUBJECT") & "</td>" & strLE & _
						"<td>archive schedule: " & "<input type=""text"" name=""archSched" & Trim(drs("FORUM_ID")) & """ size=""3"" value=""" & drs("F_ARCHIVE_SCHED") & """ maxlength=""3""> days" & "</span></td>" & strLE & _
						"</tr>" & strLE
					thisCat = drs("CAT_ID")
					drs.movenext
				loop
				Response.Write "<tr>" & strLE & _
					"<td class=""c"" colspan=""2""><input type=""submit"" name=""submit1"" value=""Update Schedule""></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"</form>" & strLE
			end if
			set drs = nothing
			Response.Write "<p class=""c""><a href=""admin_forums.asp"">Back to Forums Administration</a></p>" & strLE
		end if
end Select
Call WriteFooter
Response.End
Function GetForumName(fID)
	'## Forum_SQL
	strSql = "SELECT F.F_SUBJECT FROM " & strTablePrefix & "FORUM F WHERE F.FORUM_ID = " & fID
	set rsGetForumName = my_Conn.Execute(strSql)
	if rsGetForumName.bof or rsGetForumName.eof then GetForumName = "" else GetForumName = rsGetForumName("F_SUBJECT")
	set rsGetForumName = nothing
end Function
%>
