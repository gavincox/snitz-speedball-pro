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
%><!--#INCLUDE FILE="config.asp"--><%
'## Do Cookie stuffs with reload
nRefreshTime = Request.Cookies(strCookieURL & "Reload")
if Request.form("cookie") = "1" then
	if strSetCookieToForum = 1 then Response.Cookies(strCookieURL & "Reload").Path = strCookieURL
	Response.Cookies(strCookieURL & "Reload") = Request.Form("RefreshTime")
	Response.Cookies(strCookieURL & "Reload").expires = strForumTimeAdjust + 365
	nRefreshTime = Request.Form("RefreshTime")
end if
if nRefreshTime = "" then nRefreshTime = 0
ActiveSince = Request.Cookies(strCookieURL & "ActiveSince")
'## Do Cookie stuffs with show last date
if Request.form("cookie") = "2" then
	ActiveSince = Request.Form("ShowSinceDateTime")
	if strSetCookieToForum = 1 then Response.Cookies(strCookieURL & "ActiveSince").Path = strCookieURL
	Response.Cookies(strCookieURL & "ActiveSince") = ActiveSince
end if
Dim ModerateAllowed
Dim HasHigherSub
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
Dim canView
HasHigherSub = false
%>
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<%
Select Case ActiveSince
	Case "LastVisit" : lastDate     = ""
	Case "LastFifteen" : lastDate   = DateToStr(DateAdd("n",-15,strForumTimeAdjust))
	Case "LastThirty" : lastDate    = DateToStr(DateAdd("n",-30,strForumTimeAdjust))
	Case "LastFortyFive" : lastDate = DateToStr(DateAdd("n",-45,strForumTimeAdjust))
	Case "LastHour" : lastDate      = DateToStr(DateAdd("h",-1,strForumTimeAdjust))
	Case "Last2Hours" : lastDate    = DateToStr(DateAdd("h",-2,strForumTimeAdjust))
	Case "Last6Hours" : lastDate    = DateToStr(DateAdd("h",-6,strForumTimeAdjust))
	Case "Last12Hours" : lastDate   = DateToStr(DateAdd("h",-12,strForumTimeAdjust))
	Case "LastDay" : lastDate       = DateToStr(DateAdd("d",-1,strForumTimeAdjust))
	Case "Last2Days" : lastDate     = DateToStr(DateAdd("d",-2,strForumTimeAdjust))
	Case "LastWeek" : lastDate      = DateToStr(DateAdd("ww",-1,strForumTimeAdjust))
	Case "Last2Weeks" : lastDate    = DateToStr(DateAdd("ww",-2,strForumTimeAdjust))
	Case "LastMonth" : lastDate     = DateToStr(DateAdd("m",-1,strForumTimeAdjust))
	Case "Last2Months" : lastDate   = DateToStr(DateAdd("m",-2,strForumTimeAdjust))
	Case Else : lastDate = ""
End Select
strLE = vbNewLine
%>
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="cb/active_cb.asp" -->
<%
' Sets up the Tree structure at the top of the page
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs w50"">" & strLE & _
	"<form name=""LastDateFrm"" action=""active.asp"" method=""post"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;" & _
	"Active Topics Since " & strLE & _
	"<select name=""ShowSinceDateTime"" size=""1"" onchange=""SetLastDate();"">" & strLE & _
	"<option value=""LastVisit"""
if ActiveSince = "LastVisit" or ActiveSince = "" then Response.Write " selected"
Response.Write ">&nbsp;Last Visit on " & ChkDate(Session(strCookieURL & "last_here_date"),"",true) & "&nbsp;</option>" & strLE & _
	"<option value=""LastFifteen""" & chkSelect(ActiveSince,"LastFifteen") & ">&nbsp;Last 15 minutes</option>" & strLE & _
	"<option value=""LastThirty""" & chkSelect(ActiveSince,"LastThirty") & ">&nbsp;Last 30 minutes</option>" & strLE & _
	"<option value=""LastFortyFive""" & chkSelect(ActiveSince,"LastFortyFive") & ">&nbsp;Last 45 minutes</option>" & strLE & _
	"<option value=""LastHour""" & chkSelect(ActiveSince,"LastHour") & ">&nbsp;Last Hour</option>" & strLE & _
	"<option value=""Last2Hours""" & chkSelect(ActiveSince,"Last2Hours") & ">&nbsp;Last 2 Hours</option>" & strLE & _
	"<option value=""Last6Hours""" & chkSelect(ActiveSince,"Last6Hours") & ">&nbsp;Last 6 Hours</option>" & strLE & _
	"<option value=""Last12Hours""" & chkSelect(ActiveSince,"Last12Hours") & ">&nbsp;Last 12 Hours</option>" & strLE & _
	"<option value=""LastDay""" & chkSelect(ActiveSince,"LastDay") & ">&nbsp;Yesterday</option>" & strLE & _
	"<option value=""Last2Days""" & chkSelect(ActiveSince,"Last2Days") & ">&nbsp;Last 2 Days</option>" & strLE & _
	"<option value=""LastWeek""" & chkSelect(ActiveSince,"LastWeek") & ">&nbsp;Last Week</option>" & strLE & _
	"<option value=""Last2Weeks""" & chkSelect(ActiveSince,"Last2Weeks") & ">&nbsp;Last 2 Weeks</option>" & strLE & _
	"<option value=""LastMonth""" & chkSelect(ActiveSince,"LastMonth") & ">&nbsp;Last Month</option>" & strLE & _
	"<option value=""Last2Months""" & chkSelect(ActiveSince,"Last2Months") & ">&nbsp;Last 2 Months</option>" & strLE & _
	"</select>" & strLE & _
	"<input type=""hidden"" name=""Cookie"" value=""2"">" & strLE & _
	"</form>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""filter w50"">" & strLE & _
	"<form name=""ReloadFrm"" action=""active.asp"" method=""post"">" & strLE & _
	"<input type=""hidden"" name=""Cookie"" value=""1"">" & strLE & _
	"<select name=""RefreshTime"" size=""1"" onchange=""autoReload();"">" & strLE & _
	"<option value=""0""" & chkSelect(nRefreshTime,0) & ">Don't reload automatically</option>" & strLE & _
	"<option value=""1""" & chkSelect(nRefreshTime,1) & ">Reload page every minute</option>" & strLE & _
	"<option value=""2""" & chkSelect(nRefreshTime,2) & ">Reload page every 2 minutes</option>" & strLE & _
	"<option value=""5""" & chkSelect(nRefreshTime,5) & ">Reload page every 5 minutes</option>" & strLE & _
	"<option value=""10""" & chkSelect(nRefreshTime,10) & ">Reload page every 10 minutes</option>" & strLE & _
	"<option value=""15""" & chkSelect(nRefreshTime,15) & ">Reload page every 15 minutes</option>" & strLE & _
	"</select>" & strLE & _
	"</form>" & strLE & _
	"</div>" & strLE & _
	"<div class=""maxpages""><br class=""ffs""></div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE & _
	"<table id=""content"">" & strLE & _
	"<tr>" & strLE & _
	"<th>" & strLE
If recActiveTopicsCount <> "" and (mLev > 0) then
	Response.Write "<form name=""MarkRead"" action=""active.asp"" method=""post"" style=""display:inline"">" & strLE & _
		"<input type=""hidden"" name=""AllRead"" value=""Y"">" & strLE & _
		"<input type=""hidden"" name=""BuildTime"" value=""" & DateToStr(strForumTimeAdjust) & """>" & strLE & _
		"<input type=""hidden"" name=""Cookie"" value=""2"">" & strLE & _
		"<input type=""image"" src=""" & strImageUrl & "icon_topic_all_read.gif"" alt=""Mark all topics as read"" value=""Mark all read"" id=""submit1"" title=""Mark all topics as read"">" & strLE & _
		"</form></th>" & strLE
else
	Response.Write "&nbsp;</th>" & strLE
end if
Response.Write "<th>Topic</th>" & strLE & _
	"<th>Author</th>" & strLE & _
	"<th>Replies</th>" & strLE & _
	"<th>Read</th>" & strLE & _
	"<th>Last Post</th>" & strLE
if (mlev > 0) or (lcase(strNoCookies) = "1") then
	Response.Write "<th>"
	if (mLev = 4 or mLev = 3) or (lcase(strNoCookies) = "1") then
		if UnModeratedPosts > 0 then
			UnModeratedFPosts = 0
			Response.Write "<a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","") & "</a>"
		else
			Response.Write("&nbsp;")
		end if
	else
		Response.Write("&nbsp;")
	end if
	Response.Write "</th>" & strLE
end if
Response.Write "</tr>" & strLE
if recActiveTopicsCount = "" then
	Response.Write "<tr>" & strLE & _
		"<td class=""fcc"" colspan=""6""><b>No Active Topics Found</b></td>" & strLE & _
		"</tr>" & strLE
else
	currForum              = 0
	fDisplayCount          = 0
	canAccess              = 0
	fFORUM_ID              = 0
	fF_SUBJECT             = 1
	fF_SUBSCRIPTION        = 2
	fF_STATUS              = 3
	fCAT_ID                = 4
	fCAT_NAME              = 5
	fCAT_SUBSCRIPTION      = 6
	fCAT_STATUS            = 7
	fT_STATUS              = 8
	fT_VIEW_COUNT          = 9
	fTOPIC_ID              = 10
	fT_SUBJECT             = 11
	fT_AUTHOR              = 12
	fT_REPLIES             = 13
	fT_UREPLIES            = 14
	fM_NAME                = 15
	fT_LAST_POST_AUTHOR    = 16
	fT_LAST_POST           = 17
	fT_LAST_POST_REPLY_ID  = 18
	fLAST_POST_AUTHOR_NAME = 19
	fF_PRIVATEFORUMS       = 20
	fF_PASSWORD_NEW        = 21
	for RowCount = 0 to recActiveTopicsCount
		'## Store all the recordvalues in variables first.
		Forum_ID                    = allActiveTopics(fFORUM_ID,RowCount)
		Forum_Subject               = allActiveTopics(fF_SUBJECT,RowCount)
		ForumSubscription           = allActiveTopics(fF_SUBSCRIPTION,RowCount)
		Forum_Status                = allActiveTopics(fF_STATUS,RowCount)
		Cat_ID                      = allActiveTopics(fCAT_ID,RowCount)
		Cat_Name                    = allActiveTopics(fCAT_NAME,RowCount)
		CatSubscription             = allActiveTopics(fCAT_SUBSCRIPTION,RowCount)
		Cat_Status                  = allActiveTopics(fCAT_STATUS,RowCount)
		Topic_Status                = allActiveTopics(fT_STATUS,RowCount)
		Topic_View_Count            = allActiveTopics(fT_VIEW_COUNT,RowCount)
		Topic_ID                    = allActiveTopics(fTOPIC_ID,RowCount)
		Topic_Subject               = allActiveTopics(fT_SUBJECT,RowCount)
		Topic_Author                = allActiveTopics(fT_AUTHOR,RowCount)
		Topic_Replies               = allActiveTopics(fT_REPLIES,RowCount)
		Topic_UReplies              = allActiveTopics(fT_UREPLIES,RowCount)
		Member_Name                 = allActiveTopics(fM_NAME,RowCount)
		Topic_Last_Post_Author      = allActiveTopics(fT_LAST_POST_AUTHOR,RowCount)
		Topic_Last_Post             = allActiveTopics(fT_LAST_POST,RowCount)
		Topic_Last_Post_Reply_ID    = allActiveTopics(fT_LAST_POST_REPLY_ID,RowCount)
		Topic_Last_Post_Author_Name = chkString(allActiveTopics(fLAST_POST_AUTHOR_NAME,RowCount),"display")
		Forum_PrivateForums         = allActiveTopics(fF_PRIVATEFORUMS,RowCount)
		Forum_FPasswordNew          = allActiveTopics(fF_PASSWORD_NEW,RowCount)
		if mLev = 4 then
			ModerateAllowed = "Y"
		elseif mLev = 3 and ModOfForums <> "" then
			if (strAuthType = "nt") then
				if (chkForumModerator(Forum_ID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
			else
				if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") > 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
			end if
		else
			ModerateAllowed = "N"
		end if
		if ModerateAllowed = "Y" and Topic_UReplies > 0 then Topic_Replies = Topic_Replies + Topic_UReplies
		fDisplayCount = fDisplayCount + 1
		' -- Display forum name
		if currForum <> Forum_ID then
			Response.Write "<tr class=""cathd"">" & strLE & _
				"<td colspan=""6""><a href=""default.asp?CAT_ID=" & Cat_ID & """>" & ChkString(Cat_Name,"display") & "</a>&nbsp;/&nbsp;<a href=""forum.asp?FORUM_ID=" & Forum_ID & """>" & ChkString(Forum_Subject,"display") & "</a></td>" & strLE
			if (mlev > 0) or (lcase(strNoCookies) = "1") then
				Response.Write "<td>" & strLE
				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
					ForumAdminOptions
				else
					if Cat_Status <> 0 and Forum_Status <> 0 then Call ForumMemberOptions else Response.Write "&nbsp;" & strLE
				end if
				Response.Write "</td>" & strLE
			elseif (mLev = 3) then
				Response.Write "<td>&nbsp;</td>" & strLE
			end if
			Response.Write "</tr>" & strLE
		end if
		Response.Write "<tr class=""forumrow"">" & strLE & _
			"<td>"
		' -- Set up a link to the topic and display the icon appropriate to the status of the post.
		Response.Write "<a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>"
		' - If status = 0, topic/forum/category is locked.  If status > 2, posts are unmoderated/rejected
		if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
			' DEM --> Added code for topic moderation
			if Topic_Status = 2 then
				UnApprovedFound = "Y"
				Response.Write 	getCurrentIcon(strIconFolderUnmoderated,"Topic Not Moderated","") & "</a>" & strLE
			elseif Topic_Status = 3 then
				HeldFound = "Y"
				Response.Write 	getCurrentIcon(strIconFolderHold,"Topic on Hold","") & "</a>" & strLE
				' DEM --> end of code Added for topic moderation
			elseif lcase(strHotTopic) = "1" and Topic_Replies >= intHotTopicNum then
				Response.Write getCurrentIcon(strIconFolderNewHot,"Hot Topic with New Posts","") & "</a>" & strLE
			elseif Topic_Last_Post < lastdate then
				Response.Write getCurrentIcon(strIconFolder,"No New Posts","") & "</a>" & strLE
			else
				Response.Write getCurrentIcon(strIconFolderNew,"New Posts","") & "</a>" & strLE
			end if
		else
			if Cat_Status = 0 then
				strAltText = "Category locked"
			elseif Forum_Status = 0 then
				strAltText = "Forum locked"
			else
				strAltText = "Topic locked"
			end if
			if Topic_Last_Post < lastdate then
				Response.Write getCurrentIcon(strIconFolderLocked,strAltText,"")
			else
				Response.Write getCurrentIcon(strIconFolderNewLocked,strAltText,"")
			end if
			Response.Write "</a>" & strLE
		end if
		Response.Write "</td>" & strLE & _
			"<td class=""fdetail"">" & strLE & _
			"<span class=""smt""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkString(Topic_Subject,"title") & "</a></span>&nbsp;" & strLE
		if strShowPaging = "1" then TopicPaging()
		Response.Write "</td>" & strLE & _
			"<td><span class=""smt"">" & profileLink(chkString(Member_Name,"display"),Topic_Author) & "</span></td>" & strLE & _
			"<td>" & Topic_Replies & "</td>" & strLE & _
			"<td>" & Topic_View_Count & "</td>" & strLE
		if IsNull(Topic_Last_Post_Author) then
			strLastAuthor = ""
		else
			strLastAuthor = "<br>by: <span class=""smt"">" & profileLink(Topic_Last_Post_Author_Name,Topic_Last_Post_Author) & "</span>"
			if strJumpLastPost = "1" then strLastAuthor = strLastAuthor & "&nbsp;" & DoLastPostLink
		end if
		Response.Write "<td class=""flastpost""><b>" & ChkDate(Topic_Last_Post, "</b>&nbsp;" ,true) & strLastAuthor & "</td>" & strLE
		if (mlev > 0) or (lcase(strNoCookies) = "1") then
			Response.Write "<td class=""options"">" & strLE
			if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
				call TopicAdminOptions
			else
				if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
					call TopicMemberOptions
				else
					Response.Write "&nbsp;" & strLE
				end if
			end if
			Response.Write "</td>" & strLE
		elseif (mLev = 3) then
			Response.Write "<td>&nbsp;</td>" & strLE
		end if
		Response.Write "</tr>" & strLE
		currForum = Forum_ID
	next
	if fDisplayCount = 0 then
		Response.Write "<tr class=""forumrow"">" & strLE & _
			"<td colspan=""" & aGetColspan(6,5) & """><b>No Active Topics Found</b></td>" & strLE & _
			"</tr>" & strLE
	end if
end if
Response.Write "</table>" & strLE & _
	"<!-- /content -->" & strLE & strLE & _
	"<div id=""post-content"">" & strLE & _
	"<div class=""maxpages""><br style=""font-size: 6px;""></div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"<div class=""fkey w50"">" & strLE & _
	getCurrentIcon(strIconFolderNew,"New Posts","class=""vam""") & " New posts since last logon<br>" & strLE & _
	getCurrentIcon(strIconFolder,"Old Posts","class=""vam""") & " Old Posts"
if lcase(strHotTopic) = "1" then Response.Write " (" & getCurrentIcon(strIconFolderHot,"Hot Topic","class=""vam""") & "&nbsp;" & intHotTopicNum & " replies or more)<br>" & strLE
Response.Write getCurrentIcon(strIconFolderLocked,"Locked Topic","class=""vam""") & " Locked topic<br>" & strLE
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then Response.Write getCurrentIcon(strIconFolderHold,"Held Topic","class=""vam""") & " Held Topic<br>" & strLE
if UnapprovedFound = "Y" then Response.Write getCurrentIcon(strIconFolderUnmoderated,"UnModerated Topic","class=""vam""") & " UnModerated Topic<br>" & strLE
' DEM --> End of Code added for moderation
Response.Write "</div>" & strLE & _
	"<!-- /fkey -->" & strLE & _
	"<div class=""jumpto w50"">" & strLE
%><!--#INCLUDE FILE="inc_jump_to.asp" --><%
Response.Write "</div>" & strLE & _
	"<!-- /jumpto -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /post-content -->" & strLE & strLE & _
	"<script>" & strLE & _
	"<!-- " & vbNewLine & _
	"if (document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value > 0) {" & strLE & _
	"reloadTime = 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value" & strLE & _
	"self.setInterval('autoReload()', 60000 * document.ReloadFrm.RefreshTime.options[document.ReloadFrm.RefreshTime.selectedIndex].value)" & strLE & _
	"}" & strLE & _
	"// -->" & strLE & _
	"</script>" & strLE
Call WriteFooter
Response.End
%>