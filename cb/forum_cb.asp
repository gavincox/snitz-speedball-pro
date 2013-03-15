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

nDays = Request.Cookies(strCookieURL & "NumDays")

if Request.form("cookie") = 1 then
	if strSetCookieToForum = "1" then
		Response.Cookies(strCookieURL & "NumDays").Path = strCookieURL
	end if
	Response.Cookies(strCookieURL & "NumDays") = Request.Form("days")
	Response.Cookies(strCookieURL & "NumDays").expires = dateAdd("yyyy", 1, strForumTimeAdjust)
	nDays  = Request.Form("Days")
	mypage = 1
end if

if request("ARCHIVE") = "true" then
	nDays = "0"
end if

if mLev = 4 then
	AdminAllowed        = 1
	ForumChkSkipAllowed = 1
elseif mLev = 3 then
	if chkForumModerator(Forum_ID, chkString(strDBNTUserName,"decode")) = "1" then
		AdminAllowed        = 1
		ForumChkSkipAllowed = 1
	else
		if lcase(strNoCookies) = "1" then
			AdminAllowed        = 1
			ForumChkSkipAllowed = 0
		else
			AdminAllowed        = 0
			ForumChkSkipAllowed = 0
		end if
	end if
elseif lcase(strNoCookies) = "1" then
	AdminAllowed        = 1
	ForumChkSkipAllowed = 0
else
	AdminAllowed        = 0
	ForumChkSkipAllowed = 0
end if

if strPrivateForums = "1" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") and ForumChkSkipAllowed = 0 then
	result = ChkForumAccess(Forum_ID, MemberID, true)
end if

if strModeration = "1" and AdminAllowed = 1 then
	UnModeratedPosts = CheckForUnModeratedPosts("FORUM", Cat_ID, Forum_ID, 0)
end if

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs, strTopicSubs
if MySubCount > 0 then
	strSubString = PullSubscriptions(0, 0, 0)
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
end if

'## Forum_SQL - Find out the Category/Forum status and if it Exists
strSql = "SELECT C.CAT_STATUS, C.CAT_SUBSCRIPTION, " & _
	 "C.CAT_MODERATION, C.CAT_NAME, C.CAT_ID, " & _
	 "F.F_STATUS, F.F_SUBSCRIPTION, " & _
	 "F.F_MODERATION, F_DEFAULTDAYS, F.F_SUBJECT " & _
	 " FROM " & strTablePrefix & "CATEGORY C, " & _
	 strTablePrefix & "FORUM F " & _
	 " WHERE F.FORUM_ID = " & Forum_ID & _
	 " AND C.CAT_ID = F.CAT_ID " & _
	 " AND F.F_TYPE = 0"

set rsCFStatus = my_Conn.Execute (StrSql)

if rsCFStatus.EOF or rsCFStatus.BOF then
	rsCFStatus.close
	set rsCFStatus = nothing
	Response.Redirect("default.asp")
else
	Cat_ID             = rsCFStatus("CAT_ID")
	Cat_Name           = rsCFStatus("CAT_NAME")
	Cat_Status         = rsCFStatus("CAT_STATUS")
	Cat_Subscription   = rsCFStatus("CAT_SUBSCRIPTION")
	Cat_Moderation     = rsCFStatus("CAT_MODERATION")
	Forum_Status       = rsCFStatus("F_STATUS")
	Forum_Subject      = rsCFStatus("F_SUBJECT")
	Forum_Subscription = rsCFStatus("F_SUBSCRIPTION")
	Forum_Moderation   = rsCFStatus("F_MODERATION")
	if nDays = "" then
		nDays = rsCFStatus("F_DEFAULTDAYS")
	end if
	rsCFStatus.close
	set rsCFStatus = nothing
end if

if strModeration = 1 and Cat_Moderation = 1 and (Forum_Moderation = 1 or Forum_Moderation = 2) then
	Moderation = "Y"
end if
' DEM --> End of Code added for Moderation

if nDays = "" then
	nDays = 30
end if

defDate = DateToStr(dateadd("d",-(nDays),strForumTimeAdjust))

'## Forum_SQL - Get all topics from DB
strSql = "SELECT T.T_STATUS, T.CAT_ID, T.FORUM_ID, T.TOPIC_ID, T.T_VIEW_COUNT, T.T_SUBJECT, "
strSql = strSql & "T.T_AUTHOR, T.T_STICKY, T.T_REPLIES, T.T_UREPLIES, T.T_LAST_POST, T.T_LAST_POST_AUTHOR, "
strSql = strSql & "T.T_LAST_POST_REPLY_ID, M.M_NAME, MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "

strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS M, "
strSql2 = strSql2 & strActivePrefix & "TOPICS T, "
strSql2 = strSql2 & strMemberTablePrefix & "MEMBERS AS MEMBERS_1 "

strSql3 = " WHERE M.MEMBER_ID = T.T_AUTHOR "
strSql3 = strSql3 & " AND T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID "
strSql3 = strSql3 & " AND T.FORUM_ID = " & Forum_ID & " "
if nDays = "-1" then
	if strStickyTopic = "1" then
		strSql3 = strSql3 & " AND (T.T_STATUS <> 0 OR T.T_STICKY = 1)"
	else
		strSql3 = strSql3 & " AND T.T_STATUS <> 0 "
	end if
end if
if nDays > "0" then
	if strStickyTopic = "1" then
		strSql3 = strSql3 & " AND (T.T_LAST_POST > '" & defDate & "' OR T.T_STICKY = 1)"
	else
		strSql3 = strSql3 & " AND T.T_LAST_POST > '" & defDate & "'"
	end if
end if
' DEM --> if not a Moderator, all unapproved posts should not be viewed.
if AdminAllowed = 0 then
	strSql3 = strSql3 & " AND ((T.T_AUTHOR <> " & MemberID
	strSql3 = strSql3 & " AND T.T_STATUS < "  ' Ignore unapproved/rejected posts
	if Moderation = "Y" then
		strSql3 = strSql3 & "2"  ' Ignore unapproved posts
	else
		strSql3 = strSql3 & "3"  ' Ignore any hold posts
	end if
	strSql3 = strSql3 & ") OR T.T_AUTHOR = " & MemberID & ")"
end if

strSql4 = " ORDER BY"
if strStickyTopic = "1" then
	strSql4 = strSql4 & " T.T_STICKY DESC, "
end if
if strtopicsortfld = "author" then
	strSql4 = strSql4 & " M." & strSortCol & " "
else
	strSql4 = strSql4 & " T." & strSortCol & " "
end if

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		intOffset = cLng((mypage-1) * strPageSize)
		strSql5   = strSql5 & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(TOPIC_ID) AS PAGECOUNT "

	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing

	If iPageTotal > 0 then
		inttotaltopics = iPageTotal
		maxpages = (iPageTotal \ strPageSize )
		if iPageTotal mod strPageSize <> 0 then
			maxpages = maxpages + 1
		end if
		if iPageTotal < (strPageSize + 1) then
			intGetRows = iPageTotal
		elseif (mypage * strPageSize) > iPageTotal then
			intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
		else
			intGetRows = strPageSize
		end if
	else
		iPageTotal     = 0
		inttotaltopics = iPageTotal
		maxpages       = 0
	end if

	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			arrTopicData = rs.GetRows(intGetRows)
			iTopicCount = UBound(arrTopicData, 2)
		rs.close
		set rs = nothing
	else
		iTopicCount = ""
	end if

else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		if not rs.EOF then
			rs.movefirst
			rs.pagesize     = strPageSize
			inttotaltopics  = cLng(rs.recordcount)
			rs.absolutepage = mypage '**
			maxpages        = cLng(rs.pagecount)
			arrTopicData    = rs.GetRows(strPageSize)
			iTopicCount     = UBound(arrTopicData, 2)
		else
			iTopicCount = ""
			inttotaltopics = 0
		end if
	rs.Close
	set rs = nothing
end if


Sub PostNewTopic()
	if Cat_Status = 0 or Forum_Status = 0 then
		if (AdminAllowed = 1) then
			Response.Write "<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderLocked,"Category Locked","class=""vam""") & "</a>&nbsp;<span class=""label""><a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>New Topic</a></span><br>" & strLE
		else
			Response.Write getCurrentIcon(strIconFolderLocked,"Category Locked","class=""vam""") & "<span class=""label"">&nbsp;Category Locked</span><br>" & strLE
		end if
	else
		if Forum_Status <> 0 then
			Response.Write "<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>New Topic</a></span><br>" & strLE
		else
		    Response.Write getCurrentIcon(strIconFolderLocked,"Forum Locked","class=""vam""") & "<span class=""label"">&nbsp;Forum Locked</span><br>" & strLE
		end if
	end if
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription < 4 and strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription = 1 and (mlev > 0) and strEmail = 1 then
		if InArray(strForumSubs, Forum_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, 0, "Y") & strLE
		elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, 0, "Y") & strLE
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub

sub ForumAdminOptions()
	if (AdminAllowed = 1) then
		if Cat_Status = 0 then
			if mlev = 4 then
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & strLE
			else
				Response.Write "" & getCurrentIcon(strIconFolderLocked,"Category Locked","") & strLE
			end if
		else
			if Forum_Status <> 0 then
				Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Forum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Forum","") & "</a>" & strLE
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & strLE
			end if
		end if
		if (Cat_Status <> 0 and Forum_Status <> 0) or (AdminAllowed = 1) then
			Response.Write "<a href=""post.asp?method=EditForum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "&amp;type=0"">" & getCurrentIcon(strIconFolderPencil,"Edit Forum Properties","") & "</a>" & strLE
		end if
		if mLev = 4 or (lcase(strNoCookies) = "1") then
			Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Forum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Forum","") & "</a>" & strLE
			if strArchiveState = "1" then Response.Write("<a href=""admin_forums.asp?action=archive&amp;id=" & Forum_ID & """>" & getCurrentIcon(strIconFolderArchive,"Archive Forum","") & "</a>" & strLE)
		end if
		' DEM --> Start of Code for Moderated Posting
        if (UnModeratedPosts > 0) and (AdminAllowed = 1) then
			Response.Write "<a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","") & "</a>" & strLE
        end if
    	' DEM --> End of Code for Moderated Posting
	end if
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write "<form name=""PageNum" & fnum & """ action=""forum.asp"">" & strLE & _
			"<input name=""FORUM_ID"" type=""hidden"" value=""" & Forum_ID & """>" & strLE & _
			"<input name=""sortfield"" type=""hidden"" value=""" & strtopicsortfld & """>" & strLE & _
			"<input name=""sortorder"" type=""hidden"" value=""" & strtopicsortord & """>" & strLE
		if ArchiveView = "true" then Response.write "<input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & strLE
		if fnum = 1 then
			Response.Write "<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE
		else
			Response.Write "<b>There are " & maxpages & " pages of topics: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		if fnum = 1 then
			Response.Write "</select><b> of " & maxPages & "</b>" & strLE
		else
			Response.Write "</select>" & strLE
		end if
		Response.Write "</form>" & strLE
	end if
end sub

sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then
		mxpages = int(mxpages) + 1
	end if
	if mxpages > 1 then
		Response.Write "<table class=""tp tnb"">" & strLE
		Response.Write "<tr>" & strLE
		Response.Write "<td>" & getCurrentIcon(strIconPosticon,"Pages:","class=""vam""") & "</td>" & strLE
		for counter = 1 to mxpages
			ref = "<td>"
			'if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then
			'	ref = ref & "&nbsp;"
			'end if
			ref = ref & widenum(counter) & "<span class=""smt""><a href=""topic.asp?"
			ref = ref & ArchiveLink
		    ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&amp;whichpage=" & counter
			ref = ref & """>" & counter & "</a></span></td>"
			Response.Write ref & strLE
			if counter mod strPageNumberSize = 0 and counter < mxpages then
				Response.Write "</tr>" & strLE
				Response.Write "<tr>" & strLE
				Response.Write "<td>&nbsp;</td>" & strLE
			end if
		next
        Response.Write "</tr>" & strLE
        Response.Write "</table>" & strLE
	end if
end sub

sub TopicAdminOptions()
	if strStickyTopic = "1" then
		if Topic_Sticky then
			Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=STopic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Topic_ForumID & "')"">" & getCurrentIcon(strIconGoDown,"Make Topic Un-Sticky","") & "</a>" & strLE
		else
			Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=STopic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Topic_ForumID & "')"">" & getCurrentIcon(strIconGoUp,"Make Topic Sticky","") & "</a>" & strLE
		end if
	end if
	if Cat_Status = 0 then
		Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","") & "</a>" & strLE
	else
		if Forum_Status = 0 then
			Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Forum","") & "</a>" & strLE
		else
			if Topic_Status <> 0 then
				Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconLock,"Lock Topic","") & "</a>" & strLE
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Topic","") & "</a>" & strLE
			end if
		end if
	end if
	if (AdminAllowed = 1) or (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) then
		Response.Write "<a href=""post.asp?" & ArchiveLink & "method=EditTopic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","") & "</a>" & strLE
	end if
	Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","") & "</a>" & strLE
	if Topic_Status <= 1 and ArchiveView = "" then
		Response.Write "<a href=""post.asp?" & ArchiveLink & "method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","") & "</a>" & strLE
	end if
	' DEM --> Start of Code for Full Moderation
        if Topic_Status > 1 then
		TopicString = "TOPIC_ID=" & Topic_ID & "&amp;CAT_ID=" & Cat_ID & "&amp;FORUM_ID=" & Forum_ID
               	Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","") & "</a>" & strLE
        end if
	' DEM --> End of Code for Full Moderation
 	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N") & strLE
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N") & strLE
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub

sub TopicMemberOptions()
        if ((Topic_Status > 0 and Topic_Author = MemberID) or (AdminAllowed = 1)) and ArchiveView = "" then
		Response.Write "<a href=""post.asp?" & ArchiveLink & "method=EditTopic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconPencil,"Edit Topic","") & "</a>" & strLE
	end if
        if ((Topic_Status > 0 and Topic_Author = MemberID and Topic_Replies = 0) or (AdminAllowed = 1)) and ArchiveView = "" then
		Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Topic","") & "</a>" & strLE
	end if
	if Topic_Status <= 1 and ArchiveView = "" then
		Response.Write "<a href=""post.asp?" & ArchiveLink & "method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","") & "</a>" & strLE
 	end if
	' DEM --> Start of Code added to handle subscription processing.
	if (strSubscription > 0) and (Cat_Subscription > 0) and Forum_Subscription > 0 and strEmail = 1 then
		if InArray(strTopicSubs, Topic_ID) then
			Response.Write ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "N") & strLE
		elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
			Response.Write ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "N") & strLE
		end if
	end if
	' DEM --> End of code added to handle subscription processing.
end sub

Function DoLastPostLink()
	if Topic_Replies < 1 or Topic_LastPostReplyID = 0 then
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	elseif Topic_LastPostReplyID <> 0 then
		PageLink = "whichpage=-1&amp;"
		AnchorLink = "&amp;REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?" & ArchiveLink & PageLink & "TOPIC_ID=" & Topic_ID & AnchorLink & Topic_LastPostReplyID & """>" & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function


function CheckSelected(chkval1, chkval2)
	if IsNumeric(chkval1) then chkval1 = cLng(chkval1)
	if (chkval1 = chkval2) then
		CheckSelected = " selected"
	else
		CheckSelected = ""
	end if
end function
%>