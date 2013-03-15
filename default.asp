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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<%
strLE = "" 'vbnewline 'compress output override
%>
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="cb/default_cb.asp" -->
<%
if Cat_ID <> "" then
	Cat_Name = allCategoryData(2,0)
	Response.Write "<script type=""text/javascript"">" & strLE & _
		"<!-- " & vbNewLine & _
		"document.title='" & chkString(Cat_Name,"pagetitle") & " - " & chkString(strForumTitle,"pagetitle") & "';" & strLE & _
		"// --></script>" & strLE
end if
Response.Write "<div id=""pre-content"">" & strLE
' If Whole Board Subscription is allowed, check for a subscription by this user.
if strSubscription = 1 and strEmail = 1 and strDBNTUserName <> "" then
	Response.Write "<div id=""subscription"">" & strLE
	if strBoardSubs = "Y" then Response.Write ShowSubLink ("U", 0, 0, 0, "Y") Else Response.Write ShowSubLink ("S", 0, 0, 0, "Y")
	Response.Write "</div>" & strLE & _
		"<!-- /subscription -->" & strLE
end if
ShowLastHere = (mLev > 0)
if strShowStatistics <> "1" then
	Response.Write "<div id=""statistics"">" & strLE
	if ShowLasthere then Response.Write "You last visited: " & ChkDate(Session(strCookieURL & "last_here_date"), " - " ,true) & "<br>" & strLE
	Response.Write "There are " & Posts & " Posts in " & Topics & " Topics and " & Users & " Users" & strLE & _
		"</div>" & strLE & _
		"<!-- /statistics -->" & strLE
end if
Response.Write "</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE & _
	"<table id=""content"">" & strLE & _
	"<tr>" & strLE
if Cat_ID <> "" then Response.Write "<th><a href=""default.asp"">" & getCurrentIcon(strIconFolder,"Show All Categories","") & "</a>" else Response.Write "<th>&nbsp;"
Response.Write "</th>" & strLE & _
	"<th>"
if strGroupCategories = "1" then Response.Write GROUPNAME else Response.Write "Forum"
Response.Write "</th>" & strLE & _
	"<th>Topics</th>" & strLE & _
	"<th>Posts</th>" & strLE & _
	"<th>Last Post</th>" & strLE
if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write "<th>Moderator(s)</th>" & strLE
Response.Write "<th>"
if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then Call PostingOptions() else Response.write "&nbsp;"
Response.Write "</th>" & strLE & _
	"</tr>" & strLE
If recCategoryCount = "" then
	Response.Write "<tr class=""cathd"">" & strLE & _
		"<td colspan="""
	if (strShowModerators = "1") or (mlev > 0 ) then Response.Write "6" else Response.Write "5"
	Response.Write """><b>No Categories/Forums Found</b></td>" & strLE & _
		"<td>&nbsp;</td>" & strLE & _
		"</tr>" & strLE
else
	intPostCount          = 0
	intTopicCount         = 0
	intForumCount         = 0
	strLastPostDate       = ""
	cCAT_ID               = 0
	cCAT_STATUS           = 1
	cCAT_NAME             = 2
	cCAT_ORDER            = 3
	cCAT_SUBSCRIPTION     = 4
	cCAT_MODERATION       = 5
	fFORUM_ID             = 0
	fF_STATUS             = 1
	fCAT_ID               = 2
	fF_SUBJECT            = 3
	fF_URL                = 4
	fF_TOPICS             = 5
	fF_COUNT              = 6
	fF_LAST_POST          = 7
	fF_LAST_POST_TOPIC_ID = 8
	fF_LAST_POST_REPLY_ID = 9
	fF_TYPE               = 10
	fF_ORDER              = 11
	fF_A_COUNT            = 12
	fF_SUBSCRIPTION       = 13
	fF_PRIVATEFORUMS      = 14
	fF_PASSWORD_NEW       = 15
	fMEMBER_ID            = 16
	fM_NAME               = 17
	fT_REPLIES            = 18
	fT_UREPLIES           = 19
	fF_DESCRIPTION        = 20
	blnHiddenForums = false
	for iCategory = 0 to recCategoryCount
		CatID            = allCategoryData(cCAT_ID,iCategory)
		CatStatus        = allCategoryData(cCAT_STATUS,iCategory)
		CatName          = allCategoryData(cCAT_NAME,iCategory)
		CatOrder         = allCategoryData(cCAT_NAME,iCategory)
		CatSubscription  = allCategoryData(cCAT_SUBSCRIPTION,iCategory)
		CatModeration    = allCategoryData(cCAT_MODERATION,iCategory)
		chkDisplayHeader = true
		bContainsForum   = False
		if recForumCount <> "" then
			for iForumCheck = 0 to recForumCount
				if CatID = allForumData(fCAT_ID, iForumCheck) then bContainsForum = True
			next
		end if
		if (recForumCount = "" or not bContainsForum) and (mLev = 4) then
			Response.Write "<tr class=""cathd"">" & strLE & _
				"<td colspan=""" & sGetColspan(6,5) & """>"
			if Cat_ID = "" then
				Response.Write "<a href=""default.asp?CAT_ID=" & CatID & """ title=""View only the Forums in this Category""><b>" & ChkString(CatName,"display") & "</b></a></td>" & strLE
			else
				Response.Write "<b>" & ChkString(CatName,"display") & "</b></td>" & strLE
			end if
			if (mlev = 4) or (lcase(strNoCookies) = "1") then
				Response.Write "<td class=""options ccc vat""><b>"
				Call CategoryAdminOptions()
				Response.Write "</b></td>" & strLE
			end if
			Response.Write "</tr>" & strLE & _
				"<tr class=""forumrow"">" & strLE & _
				"<td>&nbsp;</td>" & strLE & _
				"<td class=""l"" colspan=""" & sGetColspan(6,5) & """><b>No Forums Found</b></td>" & strLE & _
				"</tr>" & strLE
		else
			for iForum = 0 to recForumCount
				if CatID = allForumData(fCAT_ID, iForum) then '## Forum exists
					ForumID              = allForumData(fFORUM_ID,iForum)
					ForumStatus          = allForumData(fF_STATUS,iForum)
					ForumCatID           = allForumData(fCAT_ID,iForum)
					ForumSubject         = allForumData(fF_SUBJECT,iForum)
					ForumURL             = allForumData(fF_URL,iForum)
					ForumTopics          = allForumData(fF_TOPICS,iForum)
					ForumCount           = allForumData(fF_COUNT,iForum)
					ForumLastPost        = allForumData(fF_LAST_POST,iForum)
					ForumLastPostTopicID = allForumData(fF_LAST_POST_TOPIC_ID,iForum)
					ForumLastPostReplyID = allForumData(fF_LAST_POST_REPLY_ID,iForum)
					ForumFType           = allForumData(fF_TYPE,iForum)
					ForumOrder           = allForumData(fF_ORDER,iForum)
					ForumACount          = allForumData(fF_A_COUNT,iForum)
					ForumSubscription    = allForumData(fF_SUBSCRIPTION,iForum)
					ForumPrivateForums   = allForumData(fF_PRIVATEFORUMS,iForum)
					ForumFPasswordNew    = allForumData(fF_PASSWORD_NEW,iForum)
					ForumMemberID        = allForumData(fMEMBER_ID,iForum)
					ForumMemberName      = allForumData(fM_NAME,iForum)
					ForumTopicReplies    = allForumData(fT_REPLIES,iForum)
					ForumTopicUReplies   = allForumData(fT_UREPLIES,iForum)
					ForumDescription     = allForumData(fF_DESCRIPTION,iForum)
					Dim AdminAllowed, ModerateAllowed
					if mLev = 4 then AdminAllowed = "Y" else AdminAllowed = "N"
					if mLev = 4 then
						ModerateAllowed = "Y"
					elseif mLev = 3 and ModOfForums <> "" then
						if (strAuthType = "nt") then
							if (chkForumModerator(ForumID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
						else
							if (instr("," & ModOfForums & "," ,"," & ForumID & ",") <> 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
						end if
					else
						ModerateAllowed = "N"
					end if
					if ModerateAllowed = "Y" and ForumTopicUReplies > 0 then ForumTopicReplies = ForumTopicReplies + ForumTopicUReplies
					if ChkDisplayForum(ForumPrivateForums,ForumFPasswordNew,ForumID,MemberID) then
						if ForumFType <> "1" then
							intPostCount  = intPostCount + ForumCount
							intTopicCount = intTopicCount + ForumTopics
							intForumCount = intForumCount + 1
							if ForumLastPost > strLastPostDate then
								strLastPostDate        = ForumLastPost
								intLastPostTopic_ID    = ForumLastPostTopicID
								intLastPostReply_ID    = ForumLastPostReplyID
								intTopicReplies        = ForumTopicReplies
								intLastPostForum_ID    = ForumID
								intLastPostMember_ID   = ForumMemberID
								strLastPostMember_Name = ForumMemberName
							end if
						end if
						if chkDisplayHeader then
							Call DoHideCategory(CatID)
							Response.Write "<tr class=""cathd"">" & strLE & _
								"<td colspan=""" & sGetColspan(6,5) & """>"
							'##### This code will specify whether or not to show the forums under a category #####
							HideForumCat = strUniqueID & "HideCat" & CatID
				 			if Request.Cookies(HideForumCat) = "Y" then
						        Response.Write "<a href=""" & ScriptName & "?" & HideForumCat & "=N&amp;CAT_ID=" & Cat_ID & """>" & getCurrentIcon(strIconPlus,"Expand This Category","") & "</a>"
							else
					       		Response.Write "<a href=""" & ScriptName & "?" & HideForumCat & "=Y&amp;CAT_ID=" & Cat_ID & """>" & getCurrentIcon(strIconMinus,"Collapse This Category","") & "</a>"
							end if
							if Cat_ID = "" then
								Response.Write "&nbsp;<a href=""default.asp?CAT_ID=" & CatID & """ title=""View only the Forums in this Category""><b>" & ChkString(CatName,"display") & "</b></a>" & strLE
							else
								Response.Write 	"&nbsp;<b>" & ChkString(CatName,"display") & "</b>" & strLE
							end if
							Response.Write "&nbsp;&nbsp;</td>" & strLE
							'##### Above code will specify whether or not to show the forums under a category #####
							Response.Write "<td class=""options""><b>"
							if (mLev = 4 or mLev = 3) or (lcase(strNoCookies) = "1") then
								call CategoryAdminOptions()
							elseif (mLev > 0) then
								call CategoryMemberOptions()
							else
								Response.Write("&nbsp;")
							end if
							Response.Write "</b></td>" & strLE & _
								"</tr>" & strLE
							chkDisplayHeader = false
						end if
						if Request.Cookies(HideForumCat) <> "Y" then  '##### added as part of Minimize Category Mod #####
							if ForumFType = 0 then
								Response.Write "<tr class=""forumrow"">" & strLE & _
									"<td>"
								ChkIsNew(ForumLastPost)
							else
								Response.Write "<tr class=""linkrow"">" & strLE & _
									"<td>"
								Response.Write "<a href=""" & ForumURL & """ target=""_blank"">" & getCurrentIcon(strIconUrl,"Visit " & chkString(ForumSubject,"display"),"") & "</a>"
							end if
							Response.Write "</td>" & strLE & _
								"<td class=""fdetail"""
							if ForumFType = 1 then Response.Write " colspan=""4"""
							Response.Write "><a href="""
							if ForumFType = 0 then
								Response.Write "forum.asp?FORUM_ID=" & ForumID
							else
								Response.Write ForumURL & """ target=""_blank"
							end if
							Response.Write """>" & chkString(ForumSubject,"display") & "</a><br>" & _
								"<span class=""fdescr"">" & formatStr(ForumDescription) & "</span></td>" & strLE
							if ForumFType = 0 then
								if IsNull(ForumTopics) then Response.Write "<td>0</td>" & strLE else Response.Write "<td>" & ForumTopics & "</td>" & strLE
								if IsNull(ForumCount) then Response.Write "<td>0</td>" & strLE else Response.Write "<td>" & ForumCount & "</td>" & strLE
								if IsNull(ForumMemberID) then
									strLastUser = "&nbsp;"
								else
									strLastUser = "<br>by:&nbsp;<span class=""smt"">" & profileLink(chkString(ForumMemberName,"display"),ForumMemberID) & "</span>"
									if strJumpLastPost = "1" then strLastUser = strLastUser & "&nbsp;" & DoLastPostLink(true)
								end if
								Response.Write "<td class=""flastpost""><b>" & ChkDate(ForumLastPost, "</b><br>" ,true) & strLastUser & "</td>" & strLE
							else
								'## Do Nothing
							end if
							if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then Response.Write "<td class=""fmods"">" & listForumModerators(ForumID) & "</td>" & strLE
							Response.Write "<td class=""options"">"
							if ModerateAllowed = "Y" or (lcase(strNoCookies) = "1") then call ForumAdminOptions else call ForumMemberOptions
							Response.Write "</td>" & strLE & _
								"</tr>" & strLE
						end if ' ##### Added as part of Minimize Category Mod #####
					else
						blnHiddenForums = true
					end if ' ChkDisplayForum()
				end if
			next '## Next Forum
		end if
	next '## Next Category
end if
if strShowStatistics = "1" then Call WriteStatistics
Response.Write "</table>" & strLE & _
	"<!-- /content -->" & strLE & strLE & _
	"<div id=""post-content"">" & strLE & _
	"<div class=""fkey"">" & strLE & _
	getCurrentIcon(strIconFolderNew,"New Posts","class=""vam""") & " Contains new posts since last visit<br>" & strLE & _
	getCurrentIcon(strIconFolder,"Old Posts","class=""vam""") & " No new posts since last visit" & strLE & _
	"</div>" & strLE & _
	"<!-- /fkey -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /post-content -->" & strLE
Call WriteFooter
%>
