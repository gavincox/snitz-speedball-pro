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
<%
Dim ArchiveView
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
Dim HasHigherSub
HasHigherSub = false
'#################################################################################

if (Request("FORUM_ID") = "" or IsNumeric(Request("FORUM_ID")) = False) and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") then
	Response.Redirect "default.asp"
else
	Forum_ID = cLng(Request("FORUM_ID"))
end if

'-------------------------------------------
' FORUM SORTING MOD VARIABLES
'-------------------------------------------

' Code Mod for mypage variable
dim mypage : mypage = request("whichpage")
if ((Trim(mypage) = "") or IsNumeric(mypage) = False) then mypage = 1
mypage = cLng(mypage)

' Topic Sorting Variables
dim strtopicsortord :strtopicsortord = request("sortorder")
dim strtopicsortfld :strtopicsortfld = request("sortfield")
dim strtopicsortday :strtopicsortday = request("days")
dim inttotaltopics : inttotaltopics  = 0
dim strSortCol, strSortOrd

Select Case strtopicsortord
	Case "asc"
		strSortOrd = " ASC"
	Case Else
		strSortOrd      = " DESC"
		strtopicsortord = "desc"
End Select

Select Case strtopicsortfld
	Case "topic" : strSortCol    = "T_SUBJECT" & strSortOrd
	Case "author" : strSortCol   = "M_NAME" & strSortOrd
	Case "replies" : strSortCol  = "T_REPLIES" & strSortOrd
	Case "views" : strSortCol    = "T_VIEW_COUNT" & strSortOrd
	Case "lastpost" : strSortCol = "T_LAST_POST" & strSortOrd
	Case Else
		strtopicsortfld = "lastpost"
		strSortCol      = "T_LAST_POST" & strSortOrd
End Select
strQStopicsort = "FORUM_ID=" & Forum_ID
'-------------------------------------------
if request("ARCHIVE") = "true" then
	strActivePrefix = strTablePrefix & "A_"
	ArchiveView     = "true"
	ArchiveLink     = "ARCHIVE=true&"
elseif request("ARCHIVE") <> "" then
	Response.Redirect "default.asp"
	Response.End
else
	strActivePrefix = strTablePrefix
	ArchiveView     = ""
	ArchiveLink     = ""
end if
%>

<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_chknew.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%

strLE = vbNewLine

%>
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="cb/forum_cb.asp" -->
<%

Response.Write "<div id=""pre-content"">" & strLE

Response.Write "<div class=""breadcrumbs w33"">" & strLE & _
	"<a href=""default.asp"">" & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "</a>&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""")
if Cat_Status <> 0 then
	Response.Write getCurrentIcon(strIconFolderOpen,"","class=""vam""")
else
	Response.Write getCurrentIcon(strIconFolderClosed,"","class=""vam""")
end if
Response.Write "&nbsp;<a href=""default.asp?CAT_ID=" & Cat_ID & """>" & ChkString(Cat_Name,"display") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""")
if ArchiveView = "true" then
	Response.Write getCurrentIcon(strIconFolderArchived,"","class=""vam""")
else
	if Cat_Status <> 0 and Forum_Status <> 0 then
		Response.Write getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""")
	else
		Response.Write getCurrentIcon(strIconFolderClosedTopic,"","class=""vam""")
	end if
end if
Response.Write "&nbsp;" & ChkString(Forum_Subject,"display") & "</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE

Response.Write "<div class=""actions w33"">" & strLE
Call PostNewTopic()
Response.Write "</div>" & strLE & _
	"<!-- /actions -->" & strLE & _

	"<div class=""filter w33"">" & strLE & _
	"<form action=""" & Request.ServerVariables("SCRIPT_NAME") & "?" & ChkString(Request.Querystring,"sqlstring") & """ method=""post"" name=""DaysFilter"">" & strLE & _
	"<select name=""Days"" onchange=""javascript:setDays();"">" & strLE & _
	"<option value=""0""" & CheckSelected(ndays,0) & ">Show all topics</option>" & strLE & _
	"<option value=""-1""" & CheckSelected(ndays,-1) & ">Show all open topics</option>" & strLE & _
	"<option value=""1""" & CheckSelected(ndays,1) & ">Show topics from last day</option>" & strLE & _
	"<option value=""2""" & CheckSelected(ndays,2) & ">Show topics from last 2 days</option>" & strLE & _
	"<option value=""5""" & CheckSelected(ndays,5) & ">Show topics from last 5 days</option>" & strLE & _
	"<option value=""7""" & CheckSelected(ndays,7) & ">Show topics from last 7 days</option>" & strLE & _
	"<option value=""14""" & CheckSelected(ndays,14) & ">Show topics from last 14 days</option>" & strLE & _
	"<option value=""30""" & CheckSelected(ndays,30) & ">Show topics from last 30 days</option>" & strLE & _
	"<option value=""60""" & CheckSelected(ndays,60) & ">Show topics from last 60 days</option>" & strLE & _
	"<option value=""120""" & CheckSelected(ndays,120) & ">Show topics from last 120 days</option>" & strLE & _
	"<option value=""365""" & CheckSelected(ndays,365) & ">Show topics from the last year</option>" & strLE & _
	"</select>" & strLE & _
	"<input type=""hidden"" name=""Cookie"" value=""1"">" & strLE & _
	"</form>" & strLE & _
	"</div>" & strLE & _
	"<!-- /filter -->" & strLE

Response.Write "<div class=""maxpages"">" & strLE
if maxpages > 1 then
	Call DropDownPaging(1)
else
	Response.Write "<br style=""font-size: 6px"">" & strLE
end if
Response.Write "</div>" & strLE & _
	"<!-- /maxpages -->" & strLE

Response.Write "</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE & _

	"<table id=""content"">" & strLE  & _
	"<tr>" & strLE & _
	"<th>&nbsp;</th>" & strLE & _
	"<th>Topic</th>" & strLE & _
	"<th>Author</th>" & strLE & _
	"<th>Replies</th>" & strLE & _
	"<th>Read</th>" & strLE & _
	"<th>Last Post</th>" & strLE
if mlev > 0 or (lcase(strNoCookies) = "1") then
	Response.Write  "<th>" & strLE
	if (AdminAllowed = 1) then
		call ForumAdminOptions
	else
		Response.Write  "&nbsp;" & strLE
	end if
	Response.Write  "</th>" & strLE
end if
Response.Write "</tr>" & strLE
if iTopicCount = "" then
	Response.Write "<tr class=""topicrow l""><td colspan=""7""><b>No Topics Found</b></td></tr>" & strLE
else
	tT_STATUS              = 0
	tCAT_ID                = 1
	tFORUM_ID              = 2
	tTOPIC_ID              = 3
	tT_VIEW_COUNT          = 4
	tT_SUBJECT             = 5
	tT_AUTHOR              = 6
	tT_STICKY              = 7
	tT_REPLIES             = 8
	tT_UREPLIES            = 9
	tT_LAST_POST           = 10
	tT_LAST_POST_AUTHOR    = 11
	tT_LAST_POST_REPLY_ID  = 12
	tM_NAME                = 13
	tLAST_POST_AUTHOR_NAME = 14

	rec = 1
	for iTopic = 0 to iTopicCount
		if (rec = strPageSize + 1) then exit for

		Topic_Status             = arrTopicData(tT_STATUS, iTopic)
		Topic_CatID              = arrTopicData(tCAT_ID, iTopic)
		Topic_ForumID            = arrTopicData(tFORUM_ID, iTopic)
		Topic_ID                 = arrTopicData(tTOPIC_ID, iTopic)
		Topic_ViewCount          = arrTopicData(tT_VIEW_COUNT, iTopic)
		Topic_Subject            = arrTopicData(tT_SUBJECT, iTopic)
		Topic_Author             = arrTopicData(tT_AUTHOR, iTopic)
		Topic_Sticky             = arrTopicData(tT_STICKY, iTopic)
		Topic_Replies            = arrTopicData(tT_REPLIES, iTopic)
		Topic_UReplies           = arrTopicData(tT_UREPLIES, iTopic)
		Topic_LastPost           = arrTopicData(tT_LAST_POST, iTopic)
		Topic_LastPostAuthor     = arrTopicData(tT_LAST_POST_AUTHOR, iTopic)
		Topic_LastPostReplyID    = arrTopicData(tT_LAST_POST_REPLY_ID, iTopic)
		Topic_MName              = arrTopicData(tM_NAME, iTopic)
		Topic_LastPostAuthorName = arrTopicData(tLAST_POST_AUTHOR_NAME, iTopic)

		if AdminAllowed = 1 and Topic_UReplies > 0 then Topic_Replies = Topic_Replies + Topic_UReplies

		Response.Write "<tr class=""topicrow"">" & strLE & _
			"<td><a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>"
		if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
			if Topic_Sticky and strStickyTopic = "1" then
				if Topic_LastPost > Session(strCookieURL & "last_here_date") then
					Response.Write getCurrentIcon(strIconFolderNewSticky,"New Sticky Topic","")
				else
					Response.Write getCurrentIcon(strIconFolderSticky,"Sticky Topic","")
				end if
			else
				' DEM --> Added code for topic moderation
				if Topic_Status = 2 then
					UnApprovedFound = "Y"
					Response.Write 	getCurrentIcon(strIconFolderUnmoderated,"Topic Not Moderated","")
				elseif Topic_Status = 3 then
					HeldFound = "Y"
					Response.Write 	getCurrentIcon(strIconFolderHold,"Topic on Hold","")
					' DEM --> end of code Added for topic moderation
				else
					Response.Write ChkIsNew(Topic_LastPost)
				end if
			end if
		else
			if ArchiveView <> "true" then
				if Cat_Status = 0 then
					strAltText = "Category Locked"
				elseif Forum_Status = 0 then
					strAltText = "Forum Locked"
				else
					strAltText = "Topic Locked"
				end if
			end if
			if ArchiveView = "true" then
				Response.Write getCurrentIcon(strIconFolderArchived,"Archived Topic","")
			elseif Topic_LastPost > Session(strCookieURL & "last_here_date") then
				if Topic_Sticky and strStickyTopic = "1" then
					Response.Write getCurrentIcon(strIconFolderNewStickyLocked,strAltText,"")
				else
					Response.Write getCurrentIcon(strIconFolderNewLocked,strAltText,"")
				end if
			else
				if Topic_Sticky and strStickyTopic = "1" then
					Response.Write getCurrentIcon(strIconFolderStickyLocked,strAltText,"")
				else
					Response.Write getCurrentIcon(strIconFolderLocked,strAltText,"")
				end if
			end if
		end if
		Response.Write "</a></td>" & strLE & _
			"<td class=""tdetail"">" & strLE
		if Topic_Sticky and strStickyTopic = "1" then Response.Write "Sticky:  "
		Response.Write "<span class=""smt""><a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & ChkString(Topic_Subject,"title") & "</a></span>&nbsp;" & strLE

		if strShowPaging = "1" then Call TopicPaging()

		Response.Write "</td>" & strLE & _
			"<td class=""tauthor"">" & profileLink(chkString(Topic_MName,"display"),Topic_Author) & "</td>" & strLE & _
			"<td class=""treplies"">" & Topic_Replies & "</td>" & strLE & _
			"<td class=""tread"">" & Topic_ViewCount & "</td>" & strLE
		if IsNull(Topic_LastPostAuthor) then
			strLastAuthor = ""
		else
			strLastAuthor = "<br>by: <span class=""smt"">" & profileLink(ChkString(Topic_LastPostAuthorName, "display"),Topic_LastPostAuthor) & "</span>"
			if (strJumpLastPost = "1") then strLastAuthor = strLastAuthor & "&nbsp;" & DoLastPostLink
		end if
		Response.Write "<td class=""tlastpost""><b>" & ChkDate(Topic_LastPost,"</b>&nbsp;",true) & strLastAuthor & "</td>" & strLE
		if mlev > 0 or (lcase(strNoCookies) = "1") then
			Response.Write "<td class=""options"">" & strLE
			if AdminAllowed = 1 then
				call TopicAdminOptions
			else
				if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
					call TopicMemberOptions
				else
					Response.Write "&nbsp;" & strLE
				end if
			end if
			Response.Write "</td>" & strLE
		end if
		Response.Write "</tr>" & strLE
		rec = rec + 1
	next
end if
'-------------------------------------------------
' TOPIC SORTING MOD
'-------------------------------------------------
Response.Write "<tr>" & strLE & _
	"<th class=""sort"" colspan=""6"">" & strLE

dim topicreclow, topicrechigh, topicpage

topicpage = mypage

if (topicpage <= 1) then topicreclow = 1 else topicreclow = ((topicpage - 1) * strPageSize) + 1

topicrechigh = topicreclow + (rec - 2)

Response.Write "<form method=""post"" name=""topicsort"" id=""pagelist"" action=""forum.asp?" & strQStopicsort & """>" & strLE
if ArchiveView = "true" then Response.Write "<input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & strLE

if inttotaltopics = 0 then
	Response.Write("No Topics Found")
elseif topicreclow = topicrechigh then
	Response.Write "Showing topic " & topicreclow & " of " & inttotaltopics
else
	Response.Write "Showing topics " & topicreclow & " - " & topicrechigh & " of " & inttotaltopics
end if

Response.Write ", sorted by: <select name=""sortfield"">" & strLE & _
	"<option value=""topic""" & CheckSelected(strtopicsortfld,"topic") & ">topic title</option>" & strLE & _
	"<option value=""author""" & CheckSelected(strtopicsortfld,"author") & ">topic author</option>" & strLE & _
	"<option value=""replies""" & CheckSelected(strtopicsortfld,"replies") & ">number of replies</option>" & strLE & _
	"<option value=""views""" & CheckSelected(strtopicsortfld,"views") & ">number of views</option>" & strLE & _
	"<option value=""lastpost""" & CheckSelected(strtopicsortfld,"lastpost") & ">last post time</option>" & strLE & _
	"</select> in <select name=""sortorder"">" & strLE & _
	"<option value=""desc""" & CheckSelected(strtopicsortord,"desc") & ">descending</option>" & strLE & _
	"<option value=""asc""" & CheckSelected(strtopicsortord,"asc") & ">ascending</option>" & strLE & _
	"</select> order, from <select name=""Days"">" & strLE & _
	"<option value=""0""" & CheckSelected(ndays,0) & ">all topics</option>" & strLE & _
	"<option value=""-1""" & CheckSelected(ndays,-1) & ">all open topics</option>" & strLE & _
	"<option value=""1""" & CheckSelected(ndays,1) & ">the last day</option>" & strLE & _
	"<option value=""2""" & CheckSelected(ndays,2) & ">the last 2 days</option>" & strLE & _
	"<option value=""5""" & CheckSelected(ndays,5) & ">the last 5 days</option>" & strLE & _
	"<option value=""7""" & CheckSelected(ndays,7) & ">the last 7 days</option>" & strLE & _
	"<option value=""14""" & CheckSelected(ndays,14) & ">the last 14 days</option>" & strLE & _
	"<option value=""30""" & CheckSelected(ndays,30) & ">the last 30 days</option>" & strLE & _
	"<option value=""60""" & CheckSelected(ndays,60) & ">the last 60 days</option>" & strLE & _
	"<option value=""90""" & CheckSelected(ndays,90) & ">the last 90 days</option>" & strLE & _
	"<option value=""120""" & CheckSelected(ndays,120) & ">the last 120 days</option>" & strLE & _
	"<option value=""365""" & CheckSelected(ndays,365) & ">the last year</option>" & strLE & _
	"</select>" & strLE & _
	"<input type=""hidden"" name=""Cookie"" value=""1""><input type=""submit"" name=""Go"" value=""Go"">" & strLE & _
	"</form>" & strLE & _
	"</th>" & strLE

if mLev > 0 or (lcase(strNoCookies) = "1") then
	Response.Write "<th class=""options"">"
	if (AdminAllowed = 1) then Call ForumAdminOptions else Response.Write "&nbsp;" & strLE
	Response.Write "</th>" & strLE
end if
Response.Write "</tr>" & strLE
'-------------------------------------------------

Response.Write "</table>" & strLE & _
	"<!-- /content -->" & strLE & strLE

Response.Write "<div id=""post-content"">" & strLE & _
	"<div class=""maxpages l"">" & strLE
	if maxpages > 1 then
		Call DropDownPaging(2)
	else
		Response.Write "<br style=""font-size: 6px;"">"
	end if
	Response.Write "</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"<div class=""tkey w33"">" & strLE & _
	getCurrentIcon(strIconFolderNew,"New Posts","class=""vam""") & " New posts since last logon<br>" & strLE & _
	getCurrentIcon(strIconFolder,"Old Posts","class=""vam""") & " Old Posts"
if lcase(strHotTopic) = "1" then Response.Write (" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","class=""vam""") & "&nbsp;" & intHotTopicNum & " replies or more)<br>" & strLE)
Response.Write getCurrentIcon(strIconFolderLocked,"Locked Topic","class=""vam""") & " Locked topic<br>" & strLE
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then Response.Write getCurrentIcon(strIconFolderHold,"Held Topic","class=""vam""") & " Held Topic<br>" & strLE
if UnapprovedFound = "Y" then Response.Write getCurrentIcon(strIconFolderUnmoderated,"UnModerated Topic","class=""vam""") & " UnModerated Topic<br>" & strLE
' DEM --> End of Code added for moderation
Response.Write "</div>" & strLE & _
	"<!-- /tkey -->" & strLE & _
	"<div class=""actions w33"">" & strLE
Call PostNewTopic()
Response.Write "</div>" & strLE & _
	"<!-- /actions -->" & strLE & _
	"<div class=""jumpto w33"">" & strLE
%><!--#INCLUDE FILE="inc_jump_to.asp" --><%
Response.Write "</div>" & strLE & _
	"<!-- /jumpto -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /post-content -->" & strLE
Call WriteFooter
Response.End
%>