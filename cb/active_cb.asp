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
if IsEmpty(Session(strCookieURL & "last_here_date")) then Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
if lastDate = "" then lastDate = Session(strCookieURL & "last_here_date")
if Request.Form("AllRead") = "Y" then
	lastDate = ChkString(Request.Form("BuildTime"),"SQLString")
	'## The redundant line below is necessary, don't delete it.
	Session(strCookieURL & "last_here_date") = lastDate
	Session(strCookieURL & "last_here_date") = lastDate
	UpdateLastHereDate lastDate,strDBNTUserName
	ActiveSince = ""
end if
if strModeration = "1" and mLev > 2 then UnModeratedPosts = CheckForUnmoderatedPosts("BOARD", 0, 0, 0)
' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
If MySubCount > 0 then
	strSubString = PullSubscriptions(0,0,0)
	strSubArray  = Split(strSubString,";")
	if uBound(strSubArray) < 0 then
		strBoardSubs = ""
		strCatSubs   = ""
		strForumSubs = ""
		strTopicSubs = ""
	else
		strBoardSubs = strSubArray(0)
		strCatSubs   = strSubArray(1)
		strForumSubs = strSubArray(2)
		strTopicSubs = strSubArray(3)
	end if
End If
if mlev = 3 then
	strSql = "SELECT FORUM_ID FROM " & strTablePrefix & "MODERATOR " & _
		" WHERE MEMBER_ID = " & MemberID
	Set rsMod = Server.CreateObject("ADODB.Recordset")
	rsMod.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsMod.EOF then
		recModCount = ""
	else
		allModData = rsMod.GetRows(adGetRowsRest)
		recModCount = UBound(allModData,2)
	end if
	RsMod.close
	set RsMod = nothing
	if recModCount <> "" then
		for x = 0 to recModCount
			if x = 0 then
				ModOfForums = allModData(0,x)
			else
				ModOfForums = ModOfForums & "," & allModData(0,x)
			end if
		next
	else
		ModOfForums = ""
	end if
else
	ModOfForums = ""
end if
if strPrivateForums = "1" and mLev < 4 then
	allAllowedForums = ""
	allowSql = "SELECT FORUM_ID, F_SUBJECT, F_PRIVATEFORUMS, F_PASSWORD_NEW"
	allowSql = allowSql & " FROM " & strTablePrefix & "FORUM"
	allowSql = allowSql & " WHERE F_TYPE = 0"
	allowSql = allowSql & " ORDER BY FORUM_ID"
	set rsAllowed = Server.CreateObject("ADODB.Recordset")
	rsAllowed.open allowSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsAllowed.EOF then
		recAllowedCount = ""
	else
		allAllowedData = rsAllowed.GetRows(adGetRowsRest)
		recAllowedCount = UBound(allAllowedData,2)
	end if
	rsAllowed.close
	set rsAllowed = nothing
	if recAllowedCount <> "" then
		fFORUM_ID        = 0
		fF_SUBJECT       = 1
		fF_PRIVATEFORUMS = 2
		fF_PASSWORD_NEW  = 3
		for RowCount = 0 to recAllowedCount
			Forum_ID            = allAllowedData(fFORUM_ID,RowCount)
			Forum_Subject       = allAllowedData(fF_SUBJECT,RowCount)
			Forum_PrivateForums = allAllowedData(fF_PRIVATEFORUMS,RowCount)
			Forum_FPasswordNew  = allAllowedData(fF_PASSWORD_NEW,RowCount)
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
			if chkForumAccessNew(Forum_PrivateForums,Forum_FPasswordNew,Forum_Subject,Forum_ID,MemberID) = true then
				if allAllowedForums = "" then
					allAllowedForums = Forum_ID
				else
					allAllowedForums = allAllowedForums & "," & Forum_ID
				end if
			end if
		next
	end if
	if allAllowedForums = "" then allAllowedForums = 0
end if
'## Forum_SQL - Get all active topics from last visit
strSql = "SELECT F.FORUM_ID, " & _
	"F.F_SUBJECT, " & _
	"F.F_SUBSCRIPTION, " & _
	"F.F_STATUS, " & _
	"C.CAT_ID, " & _
	"C.CAT_NAME, " & _
	"C.CAT_SUBSCRIPTION, " & _
	"C.CAT_STATUS, " & _
	"T.T_STATUS, " & _
	"T.T_VIEW_COUNT, " & _
	"T.TOPIC_ID, " & _
	"T.T_SUBJECT, " & _
	"T.T_AUTHOR, " & _
	"T.T_REPLIES, " & _
	"T.T_UREPLIES, " & _
	"M.M_NAME, " & _
	"T.T_LAST_POST_AUTHOR, " & _
	"T.T_LAST_POST, " & _
	"T.T_LAST_POST_REPLY_ID, " & _
	"MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME, " & _
	"F.F_PRIVATEFORUMS, " & _
	"F.F_PASSWORD_NEW " & _
	"FROM " & strMemberTablePrefix & "MEMBERS M, " & _
	strTablePrefix & "FORUM F, " & _
	strTablePrefix & "TOPICS T, " & _
	strTablePrefix & "CATEGORY C, " & _
	strMemberTablePrefix & "MEMBERS MEMBERS_1 " & _
	"WHERE T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID "
if strPrivateForums = "1" and mLev < 4 then strSql = strSql & " AND F.FORUM_ID IN (" & allAllowedForums & ") "
strSql = strSql & "AND F.F_TYPE = 0 " & _
	"AND F.FORUM_ID = T.FORUM_ID " & _
	"AND C.CAT_ID = T.CAT_ID " & _
	"AND M.MEMBER_ID = T.T_AUTHOR " & _
	"AND (T.T_LAST_POST > '" & lastDate & "'"
' DEM --> if not an admin, all unapproved posts should not be viewed.
if mlev <> 4 then
	strSql = strSql & " AND ((T.T_AUTHOR <> " & MemberID &_
		" AND T.T_STATUS < 2)"  ' Ignore unapproved/held posts
	if mlev = 3 and ModOfForums <> "" then strSql = strSql & " OR T.FORUM_ID IN (" & ModOfForums & ") "
	strSql = strSql & "  OR T.T_AUTHOR = " & MemberID & ")"
end if
if Group > 1 and strGroupCategories = "1" then
	strSql = strSql & " AND (C.CAT_ID = 0"
	if recGroupCatCount <> "" then
		for iGroupCat = 0 to recGroupCatCount
			strSql = strSql & " or C.CAT_ID = " & allGroupCatData(1, iGroupCat)
		next
		strSql = strSql & ")"
	else
		strSql = strSql & ")"
	end if
end if
strSql = strSql & ") "
strSql = strSql & " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT, T.T_LAST_POST DESC "
Set rs = Server.CreateObject("ADODB.Recordset")
if strDBType <> "mysql" then rs.cachesize = 50
rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
if rs.EOF then
	recActiveTopicsCount = ""
else
	allActiveTopics = rs.GetRows(adGetRowsRest)
	recActiveTopicsCount = UBound(allActiveTopics,2)
end if
rs.close
set rs = nothing
sub ForumAdminOptions()
	if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
		if Cat_Status = 0 then
			if mlev = 4 then
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & strLE
			else
				Response.Write "" & getCurrentIcon(strIconFolderLocked,"Category Locked","") & strLE
			end if
		else
			if Forum_Status <> 0 then
				Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Forum","") & "</a>" & strLE
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & strLE
			end if
		end if
		if (Cat_Status <> 0 and Forum_Status <> 0) or (ModerateAllowed = "Y") then
			Response.Write "<a href=""post.asp?method=EditForum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "&type=0"">" & getCurrentIcon(strIconFolderPencil,"Edit Forum Properties","class=""vam""") & "</a>" & strLE
		end if
		if mLev = 4 or lcase(strNoCookies) = "1" then Response.Write("<a href=""JavaScript:openWindow('pop_delete.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Forum","") & "</a>" & strLE)
		Response.Write "<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & strLE
		' DEM --> Start of Code added to handle subscription processing.
		if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription = 1 and strEmail = 1 then
			if InArray(strForumSubs, Forum_ID) then
				Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "N")
			elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
				Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "N")
			end if
		end if
		' DEM --> End of code added to handle subscription processing.
	end if
end sub
sub ForumMemberOptions()
	if (mlev > 0) then
		Response.Write "<a href=""post.asp?method=Topic&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & strLE
		' DEM --> Start of Code added to handle subscription processing.
		if (strSubscription > 0 and strSubscription < 4) and CatSubscription > 0 and ForumSubscription = 1 and strEmail = 1 then
			if InArray(strForumSubs, Forum_ID) then
				Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "N")
			elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
				Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "N")
			end if
		end if
	end if
end sub
sub TopicAdminOptions()
	if Cat_Status = 0 then
		Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","class=""vam""") & "</a>" & strLE
	elseif Forum_Status = 0 then
		Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Forum","class=""vam""") & "</a>" & strLE
	elseif Topic_Status <> 0 then
		Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconLock,"Lock Topic","class=""vam""") & "</a>" & strLE
	else
		Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Topic","class=""vam""") & "</a>" & strLE
	end if
	if (ModerateAllowed = "Y") or (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) then
		Response.Write "<a href=""post.asp?method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&auth=" & Topic_Author & """>" & getCurrentIcon(strIconPencil,"Edit Topic","class=""vam""") & "</a>" & strLE
	end if
	Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","class=""vam""") & "</a>" & strLE
	if Topic_Status <= 1 then
		Response.Write "<a href=""post.asp?method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","class=""vam""") & "</a>" & strLE
	end if
	' DEM --> Start of Code for Full Moderation
	if Topic_Status > 1 then
		TopicString = "TOPIC_ID=" & Topic_ID & "&CAT_ID=" & Cat_ID & "&FORUM_ID=" & Forum_ID
		Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","class=""vam""") & "</a>" & strLE
	end if
	' DEM --> End of Code for Full Moderation
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write "&nbsp;" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N")
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write "&nbsp;" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N")
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub
sub TopicMemberOptions()
	if (Topic_Status > 0 and Topic_Author = MemberID) or (ModerateAllowed = "Y") then
		Response.Write "<a href=""post.asp?method=EditTopic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","class=""vam""") & "</a>" & strLE
	end if
	if (Topic_Status > 0 and Topic_Author = MemberID and Topic_Replies = 0) or (ModerateAllowed = "Y") then
		Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?mode=Topic&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & "&CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","class=""vam""") & "</a>" & strLE
	end if
	if Topic_Status <= 1 then
		Response.Write "<a href=""post.asp?method=Reply&TOPIC_ID=" & Topic_ID & "&FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","class=""vam""") & "</a>" & strLE
	end if
	if (strSubscription < 4 and strSubscription > 0) and (CatSubscription > 0) and ForumSubscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write "&nbsp;" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N")
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write "&nbsp;" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N")
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub
sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then mxpages = int(mxpages) + 1
	if mxpages > 1 then
		Response.Write "<table class=""tp tnb"">" & strLE & _
			"<tr>" & strLE & _
			"<td>" & getCurrentIcon(strIconPosticon,"Pages:","class=""vam""") & "</td>" & strLE
		for counter = 1 to mxpages
			ref =	"<td>"
			'if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
			''	ref = ref & "&nbsp;"
			'end if
			ref = ref & widenum(counter) & "<span class=""smt""><a href=""topic.asp?"
			ref = ref & ArchiveLink
			ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&amp;whichpage=" & counter
			ref = ref & """>" & counter & "</a></span></td>"
			Response.Write ref & strLE
			if counter mod strPageNumberSize = 0 and counter < mxpages then
				Response.Write "</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td>&nbsp;</td>" & strLE
			end if
		next
		Response.Write "</tr>" & strLE & _
			"</table>" & strLE
	end if
end sub
Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_Last_Post_Reply_ID = 0 then
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	elseif Topic_Last_Post_Reply_ID <> 0 then
		PageLink       = "whichpage=-1&amp;"
		AnchorLink     = "&amp;REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Topic_Last_Post_Reply_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function
function aGetColspan(lIN, lOUT)
	if (mlev > 0 or strNoCookies = "1") then lOut = lOut + 1
	if lOut > lIn then aGetColspan = lIN else aGetColspan = lOUT
end function
%>
