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

Dim UnapprovedFound, UnModeratedPosts

if Request.QueryString("CAT_ID") <> "" and IsNumeric(Request.QueryString("CAT_ID")) = True then
	Cat_ID = cLng(Request.QueryString("CAT_ID"))
end if

scriptname = request.servervariables("script_name")

if strAutoLogon = 1 then
	if (ChkAccountReg() <> "1") then
		Response.Redirect("register.asp?mode=DoIt")
	end if
end if

if IsEmpty(Session(strCookieURL & "last_here_date")) then
	Session(strCookieURL & "last_here_date") = ReadLastHereDate(strDBNTUserName)
end if

if strModeration = "1" and mLev > 2 then
	UnModeratedPosts = CheckForUnmoderatedPosts("BOARD", 0, 0, 0)
end if

' -- Get all the high level(board, category, forum) subscriptions being held by the user
Dim strSubString, strSubArray, strBoardSubs, strCatSubs, strForumSubs
if MySubCount > 0 then
	strSubString = PullSubscriptions(0,0,0)
	strSubArray  = Split(strSubString,";")
	if uBound(strSubArray) < 0 then
		strBoardSubs = ""
		strCatSubs   = ""
		strForumSubs = ""
	else
		strBoardSubs = strSubArray(0)
		strCatSubs   = strSubArray(1)
		strForumSubs = strSubArray(2)
	end if
end If

if strShowStatistics <> "1" then
	'## Forum_SQL
	strSql = "SELECT P_COUNT, T_COUNT, U_COUNT " & _
		 " FROM " & strTablePrefix & "TOTALS"

	Set rs1 = Server.CreateObject("ADODB.Recordset")
	rs1.open strSql, my_Conn

	Users  = rs1("U_COUNT")
	Topics = rs1("T_COUNT")
	Posts  = rs1("P_COUNT")

	rs1.Close
	set rs1 = nothing
end if

if (strShowModerators = "1") or (mlev = 4 or mlev = 3) then
	'## Forum_SQL
	strSql = "SELECT MO.FORUM_ID, ME.MEMBER_ID, ME.M_NAME " & _
		 " FROM " & strTablePrefix & "MODERATOR MO" & _
		 " , " & strMemberTablePrefix & "MEMBERS ME" & _
		 " WHERE (MO.MEMBER_ID = ME.MEMBER_ID )" & _
		 " ORDER BY MO.FORUM_ID, ME.M_NAME"

	Set rsChk = Server.CreateObject("ADODB.Recordset")
	rsChk.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

	if rsChk.EOF then
		recModeratorCount = ""
	else
		allModeratorData = rsChk.GetRows(adGetRowsRest)
		recModeratorCount = UBound(allModeratorData,2)
	end if

	rsChk.close
	set rsChk = nothing

	if recModeratorCount = "" then
		fMods = "&nbsp;"
	else
		mFORUM_ID  = 0
		mMEMBER_ID = 1
		mM_NAME    = 2

			for iModerator = 0 to recModeratorCount
			ModForumID = allModeratorData(mFORUM_ID, iModerator)
			ModMemID = allModeratorData(mMEMBER_ID, iModerator)
			ModMemName = replace(allModeratorData(mM_NAME, iModerator),"|","&#124")

			if iModerator = 0 then
				strForumMods = ModForumID & "," & ModMemID & "," & ModMemName
			else
				strForumMods = strForumMods & "|" & ModForumID & "," & ModMemID & "," & ModMemName
			end if
		next
	end if
end if

'## Forum_SQL - Get all Categories from  the DB
strSql = "SELECT CAT_ID, CAT_STATUS, CAT_NAME, CAT_ORDER, CAT_SUBSCRIPTION, CAT_MODERATION " &_
	 " FROM " & strTablePrefix & "CATEGORY "
'############################## Group Cat MoD #####################################
if Cat_ID <> "" then
	strSql = strSql & " WHERE CAT_ID = " & Cat_ID
else
	if Group > 1 and strGroupCategories = "1" then
		strSql = strSql & " WHERE CAT_ID = 0"
		if recGroupCatCount <> "" then
			for iGroupCat = 0 to recGroupCatCount
				strSql = strSql & " or CAT_ID = " & allGroupCatData(1, iGroupCat)
			next
		end if
	end if
end if
'############################## Group Cat MoD #####################################
strSql = strSql & " ORDER BY CAT_ORDER ASC, CAT_NAME ASC;"

set rs = Server.CreateObject("ADODB.Recordset")
rs.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rs.EOF then
	if Cat_ID <> "" then response.redirect("default.asp")
	recCategoryCount = ""
else
	allCategoryData = rs.GetRows(adGetRowsRest)
	recCategoryCount = UBound(allCategoryData,2)
end if

rs.close
set rs = nothing

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

'## Forum_SQL - Build SQL to get forums via category
strSql = "SELECT F.FORUM_ID, F.F_STATUS, F.CAT_ID, F.F_SUBJECT, F.F_URL, F.F_TOPICS, " &_
	 "F.F_COUNT, F.F_LAST_POST, F.F_LAST_POST_TOPIC_ID, F.F_LAST_POST_REPLY_ID, F.F_TYPE, " & _
	 "F.F_ORDER, F.F_A_COUNT, F.F_SUBSCRIPTION, F_PRIVATEFORUMS, F_PASSWORD_NEW, " & _
	 "M.MEMBER_ID, M.M_NAME, " & _
         "T.T_REPLIES, T.T_UREPLIES, " & _
         "F.F_DESCRIPTION " & _
	 "FROM ((" & strTablePrefix & "FORUM F " &_
	 "LEFT JOIN " & strMemberTablePrefix & "MEMBERS M ON " &_
	 "F.F_LAST_POST_AUTHOR = M.MEMBER_ID) " & _
         "LEFT JOIN " & strTablePrefix & "TOPICS T ON " & _
         "F.F_LAST_POST_TOPIC_ID = T.TOPIC_ID) "
'############################## Group Cat MoD #####################################
if Cat_ID <> "" then
	strSql = strSql & " WHERE F.CAT_ID = " & Cat_ID
else
	if Group > 1 and strGroupCategories = "1" then
		strSql = strSql & " WHERE F.CAT_ID = 0"
		if recGroupCatCount <> "" then
			for iGroupCat = 0 to recGroupCatCount
				strSql = strSql & " OR F.CAT_ID = " & allGroupCatData(1, iGroupCat)
			next
		end if
	end if
end if
'############################## Group Cat MoD #####################################
strSql = strSql & " ORDER BY F.F_ORDER ASC, F.F_SUBJECT ASC;"
set rsForum = Server.CreateObject("ADODB.Recordset")
rsForum.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rsForum.EOF then
	recForumCount = ""
else
	allForumData  = rsForum.GetRows(adGetRowsRest)
	recForumCount = UBound(allForumData,2)
end if

rsForum.close
set rsForum = nothing

sub WriteStatistics()
	Dim Forum_Count
	Dim NewMember_Name, NewMember_Id, Member_Count
	Dim LastPostDate, LastPostLink

	Forum_Count = intForumCount
	'## Forum_SQL - Get newest membername and id from DB

	strSql = "SELECT M_NAME, MEMBER_ID FROM " & strMemberTablePrefix & "MEMBERS " &_
	" WHERE M_STATUS = 1 AND MEMBER_ID > 1 " &_
	" ORDER BY MEMBER_ID desc;"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open TopSQL(strSql,1), my_Conn
	if not rs.EOF then
		NewMember_Name = chkString(rs("M_NAME"), "display")
		NewMember_Id   = rs("MEMBER_ID")
	else
		NewMember_Name = ""
	end if
	rs.close
	set rs = nothing

	'## Forum_SQL - Get Active membercount from DB
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS WHERE M_POSTS > 0 AND M_STATUS=1"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn
	if not rs.EOF then Member_Count = rs("U_COUNT") else Member_Count = 0
	rs.close
	set rs = nothing

	'## Forum_SQL - Get membercount from DB
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS WHERE M_STATUS=1"
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn
	if not rs.EOF then User_Count = rs("U_COUNT") else User_Count = 0
	rs.close
	set rs = nothing

	LastPostDate       = ""
	LastPostLink       = ""
	LastPostAuthorLink = ""

	if not (intLastPostForum_ID = "") then
		ForumTopicReplies    = intTopicReplies
		ForumLastPostTopicID = intLastPostTopic_ID
		ForumLastPostReplyID = intLastPostReply_ID
		LastPostDate         = ChkDate(strLastPostDate,"",true)
		LastPostLink         = DoLastPostLink(false)
		LastPostAuthorLink   = " by: <span class=""smt"">" & profileLink(chkString(strLastPostMember_Name,"display"),intLastPostMember_ID) & "</span>"
	end if

	ActiveTopicCount = -1
	if not IsNull(Session(strCookieURL & "last_here_date")) then
		if not blnHiddenForums then
			'## Forum_SQL - Get ActiveTopicCount from DB
			strSql = "SELECT COUNT(" & strTablePrefix & "TOPICS.T_LAST_POST) AS NUM_ACTIVE " &_
			" FROM " & strTablePrefix & "TOPICS " &_
			" WHERE (((" & strTablePrefix & "TOPICS.T_LAST_POST)>'"& Session(strCookieURL & "last_here_date") & "'))" &_
			" AND " & strTablePrefix & "TOPICS.T_STATUS <= 1"
			set rs = Server.CreateObject("ADODB.Recordset")
			rs.open strSql, my_Conn
			if not rs.EOF then ActiveTopicCount = rs("NUM_ACTIVE") else ActiveTopicCount = 0
			rs.close
			set rs = nothing
		end if
	end if

	ArchivedPostCount = 0
	ArchivedTopicCount = 0

	if not blnHiddenForums and strArchiveState = "1" then
		'## Forum_SQL
		strSql = "SELECT P_A_COUNT, T_A_COUNT FROM " & strTablePrefix & "TOTALS"
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql, my_Conn
		if not rs.EOF then
			ArchivedPostCount  = rs("P_A_COUNT")
			ArchivedTopicCount = rs("T_A_COUNT")
		else
			ArchivedPostCount  = 0
			ArchivedTopicCount = 0
		end if
		rs.Close
		set rs = nothing
	end if

	'ShowLastHere = (cLng(chkUser(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"),-1)) > 0)
	Response.Write "<tr class=""statshd"">" & strLE & _
		"<td colspan=""" & sGetColspan(7,6) & """><b>Statistics</b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr class=""statsrow"">" & strLE & _
		"<td rowspan="""

	intStatRowSpan = 2

	if ShowLastHere then intStatRowSpan = intStatRowspan + 1
	if ArchivedPostCount > 0 and strArchiveState = "1" then intStatRowSpan = intStatRowspan + 1
	if NewMember_Name <> "" then intStatRowSpan = intStatRowSpan + 1

 	Response.Write intStatRowSpan & """>&nbsp;</td>" & strLE

	if ShowLastHere then
		Response.Write "<td colspan=""" & sGetColspan(6,5) &	""">" & _
			"You last visited on " & ChkDate(Session(strCookieURL & "last_here_date"), " " ,true) & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
		  	"<tr class=""statsrow"">" & strLE
	end if
	if intPostCount > 0 then
		Response.Write "<td colspan=""" & sGetColspan(6,5) & """>"
		if Member_Count = 1 and User_Count = 1 then
			Response.Write "1 Member has "
		else
			Response.Write Member_Count & " of " & User_Count & " <span class=""smt""><a href=""members.asp"">Members</a></span> have "
		end if
		Response.Write " made "
		if intPostCount = 1 then Response.Write "1 post " else Response.Write intPostCount & " posts"
		Response.Write " in "
		if intForumCount = 1 then Response.Write "1 forum" else Response.Write intForumCount & " forums"
		if (LastPostDate = "" or LastPostLink = "" or intPostCount = 0) then
			Response.Write "."
		else
			Response.Write ", with the last post on <span class=""smt"">" & LastPostLink & LastPostDate & "</a></span>"
			if  LastPostAuthorLink <> "" then Response.Write LastPostAuthorLink
		end if
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE & _
			"<tr class=""statsrow"">" & strLE
	end if
	Response.Write "<td colspan=""" & sGetColspan(6,5) & """>There "
	if intTopicCount = 1 then Response.Write "is" else Response.Write "are"
	Response.Write " currently "
	if intTopicCount > 0 then Response.Write intTopicCount else Response.Write "no"
	if intTopicCount = 1 then Response.Write " topic" else Response.Write " topics"
	if ActiveTopicCount > 0 then
		Response.Write " and " & ActiveTopicCount & " <span class=""smt""><a href=""active.asp"">active "
		if ActiveTopicCount = 1 then Response.Write "topic" else Response.Write "topics"
		Response.Write "</a></span> since you last visited"
	elseif blnHiddenForums and (strLastPostDate > Session(strCookieURL & "last_here_date")) and ShowLastHere then
		Response.Write " and there are <span class=""smt""><a href=""active.asp"">active topics</a></span> since you last visited"
	elseif not(ShowLastHere) then
		'Response.Write "."
	else
		Response.Write " and no active topics since you last visited"
	end if
	Response.Write "</td>" & strLE & _
		"</tr>" & strLE
	if ArchivedPostCount > 0 and strArchiveState = "1" then
		Response.Write "<tr class=""statsrow"">" & strLE & _
			"<td colspan=""" & sGetColspan(6,5) & """>There "
		if ArchivedPostCount = 1 then Response.Write "is " else Response.Write "are "
		Response.Write ArchivedPostCount & " "
		if ArchivedPostCount = 1 then Response.Write " archived post " else Response.Write " archived posts"
		if ArchivedTopicCount > 0 then
			Response.Write " in " & ArchivedTopicCount
			if ArchivedTopicCount = 1 then Response.Write " archived topic" else Response.Write " archived topics"
		end if
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE
	end if
	if NewMember_Name <> "" then
		Response.Write "<tr class=""statsrow"">" & strLE & _
			"<td colspan=""" & sGetColspan(6,5) & """>" & _
			"Please welcome our newest member: " & _
			"<span class=""smt"">" & profileLink(NewMember_Name,NewMember_Id) & "</span></td>" & strLE & _
			"</tr>" & strLE
	end if
end sub

sub PostingOptions()
	if (mlev = 4) or (lcase(strNoCookies) = "1") then

		if Session(strCookieURL & "Approval") = "15916941253" then
			Response.Write "<a href=""down.asp"">" & getCurrentIcon(strIconLock,"Shut Down the Forum","") & "</a>"
		end if

		Response.Write "&nbsp;<a href=""post.asp?method=Category"">" & getCurrentIcon(strIconFolderNewTopic,"Create New Category","") & "</a>"

		if strArchiveState = "1" then
			Response.Write "&nbsp;<a href=""admin_forums.asp"">" & getCurrentIcon(strIconFolderArchive,"Archive Forum Topics","") & "</a>"
		end if

        ' DEM --> Start of Code for Full Moderation
    	if UnModeratedPosts > 0 then
			Response.Write " <a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","") & "</a>"
			'Response.Write " <a href=""JavaScript:openWindow('pop_moderate.asp')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject all UnModerated Posts","") & "</a>"
        end if
    	' DEM --> End of Code for Full Moderation

		' DEM - Added to allow for sorting
		Response.Write "&nbsp;<a href=""Javascript:openWindow3('admin_config_order.asp')"">" & getCurrentIcon(strIconSort,"Set the order of Forums and Categories","") & "</a>"
		'############################## Group Cat MoD #####################################
		if strGroupCategories = "1" then
			Response.Write "&nbsp;<a href=""admin_config_groupcats.asp?method=Edit"">" & getCurrentIcon(strIconGroupCategories,"Configure Group Categories","") & "</a>"
		end if
		'############################## Group Cat MoD #####################################
	elseif (mlev = 3) then

        if UnModeratedPosts > 0 then
			Response.Write " <a href=""moderate.asp"">" & getCurrentIcon(strIconFolderModerate,"View All UnModerated Posts","") & "</a>"
		else
			Response.Write "&nbsp;"
		end if

	else
		Response.Write "&nbsp;"
	end if
end sub

sub ChkIsNew(dt)
	Response.Write "<a href=""forum.asp?FORUM_ID=" & ForumID & """>"
	if CatStatus <> 0 and ForumStatus <> 0 then
		if dt > Session(strCookieURL & "last_here_date") and (ForumCount > 0 or ForumTopics > 0) then
			Response.Write getCurrentIcon(strIconFolderNew,"New Posts","") & "</a>"
		else
			Response.Write getCurrentIcon(strIconFolder,"Old Posts","") & "</a>"
		end if
	elseif ForumLastPost > Session(strCookieURL & "last_here_date") then
		if CatStatus = 0 then
			strAltText = "Category Locked"
		else
			strAltText = "Forum Locked"
		end if
		Response.Write getCurrentIcon(strIconFolderNewLocked,strAltText,"") & "</a>"
	else
		if CatStatus = 0 then
			strAltText = "Category Locked"
		else
			strAltText = "Forum Locked"
		end if
		Response.Write getCurrentIcon(strIconFolderLocked,strAltText,"") & "</a>"
	end if
end sub

sub CategoryAdminOptions()
	if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then

        if (mlev = 4) or (lcase(strNoCookies) = "1") then
           	if (CatStatus <> 0) then
              	Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_lock.asp?mode=Category&amp;CAT_ID=" & CatID & "')"">" & getCurrentIcon(strIconLock,"Lock Category","") & "</a>"
           	else
           		Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&amp;CAT_ID=" & CatID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","") & "</a>"
           	end if
        end if

		if (mlev = 4) or (lcase(strNoCookies) = "1") then
			if (CatStatus <> 0) then
				Response.Write "&nbsp;<a href=""post.asp?method=EditCategory&amp;CAT_ID=" & CatID & """>" & getCurrentIcon(strIconPencil,"Edit Category Name","") & "</a>"
			end if
		end if

        if mlev = 4 or (lcase(strNoCookies) = "1") then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?mode=Category&amp;CAT_ID=" & CatID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Category","") & "</a>"
		end if

		if (mlev = 4) or (lcase(strNoCookies) = "1") then
			if (CatStatus <> 0) then
				Response.Write "&nbsp;<a href=""post.asp?method=Forum&amp;CAT_ID=" & CatID & """>" & getCurrentIcon(strIconFolderNewTopic,"Create New Forum","") & "</a>"
			end if
		end if

		if (mlev = 4) or (lcase(strNoCookies) = "1") then
			if (CatStatus <> 0) then
				Response.Write "&nbsp;<a href=""post.asp?method=URL&amp;CAT_ID=" & CatID & """>" & getCurrentIcon(strIconUrl,"Create New Web Link","") & "</a>"
			end if
		end if

		if (strSubscription = 1 or strSubscription = 2) and CatSubscription = 1 and strEmail = 1 then
			if InArray(strCatSubs,CatID) then
				Response.Write  "&nbsp;" & ShowSubLink ("U", CatID, 0, 0, "N")
			elseif strBoardSubs <> "Y" then
				Response.Write  "&nbsp;" & ShowSubLink ("S", CatID, 0, 0, "N")
			end if
		elseif mLev = "3" then
			Response.Write "&nbsp;"
		end if

	else
		Response.Write "&nbsp;"
	end if
end sub

sub CategoryMemberOptions()
	if (strSubscription = 1 or strSubscription = 2) and CatSubscription = 1 and CatStatus <> 0 and strEmail = 1 then
		if InArray(strCatSubs,CatID) then
			Response.Write  "&nbsp;" & ShowSubLink ("U", CatID, 0, 0, "N")
		elseif strBoardSubs <> "Y" then
			Response.Write  "&nbsp;" & ShowSubLink ("S", CatID, 0, 0, "N")
		end if
	else
		Response.Write "&nbsp;"
	end if
end sub

sub ForumAdminOptions()
	if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
		if ForumFType = 0 then
			if CatStatus = 0 then
				if (mlev = 4) then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&amp;CAT_ID=" & CatID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Category","") & "</a>"
				end if
			else
				if ForumStatus = 1 then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_lock.asp?mode=Forum&amp;FORUM_ID=" & ForumID & "&amp;CAT_ID=" & ForumCatID & "')"">" & getCurrentIcon(strIconLock,"Lock Forum","") & "</a>"
				else
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&amp;FORUM_ID=" & ForumID & "&amp;CAT_ID=" & ForumCatID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Forum","") & "</a>"
				end if
			end if
		end if

		if ForumFType = 0 then
			if (CatStatus <> 0 and ForumStatus <> 0) or (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
				Response.Write "&nbsp;<a href=""post.asp?method=EditForum&amp;FORUM_ID=" & ForumID & "&amp;CAT_ID=" & ForumCatID & """>" & getCurrentIcon(strIconPencil,"Edit Forum Properties","") & "</a>"
			end if
		else
			if ForumFType = 1 then
				Response.Write "&nbsp;<a href=""post.asp?method=EditURL&amp;FORUM_ID=" & ForumID & "&amp;CAT_ID=" & ForumCatID & """>" & getCurrentIcon(strIconPencil,"Edit URL Properties","") & "</a>"
			end if
		end if

		if (mlev = 4) or (lcase(strNoCookies) = "1") then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?mode=Forum&amp;FORUM_ID=" & ForumID & "&amp;CAT_ID=" & ForumCatID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Forum","") & "</a>"
		end if

		if ForumFType = 0 then
			Response.Write "&nbsp;<a href=""post.asp?method=Topic&amp;FORUM_ID=" & ForumID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>"
		end if

		if ((mlev = 4) or (lcase(strNoCookies) = "1")) and (ForumFType = 0) and (strArchiveState = "1") then
			Response.Write "&nbsp;<a href=""admin_forums.asp?action=archive&amp;id=" & ForumID & """>" & getCurrentIcon(strIconFolderArchive,"Archive Forum","") & "</a>"
		end if

		if (ForumFType = 0 and ForumACount > 0) and strArchiveState = "1" then
			Response.Write "&nbsp;<a href=""forum.asp?ARCHIVE=true&amp;FORUM_ID=" & ForumID & """>" & getCurrentIcon(strIconFolderArchived,"View Archived posts","") & "</a>"
		end if

		if (strSubscription > 0 and strSubscription < 4) and CatSubscription > 0 and ForumSubscription = 1 and strEmail = 1 then
			if InArray(strForumSubs,ForumID) then
				Response.Write "&nbsp;" & ShowSubLink ("U", ForumCatID, ForumID, 0, "N")
			elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,ForumCatID)) then
				Response.Write "&nbsp;" & ShowSubLink ("S", ForumCatID, ForumID, 0, "N")
			end if
		end if

	else
		Response.Write "&nbsp;"
	end if
end sub

sub ForumMemberOptions()
	if (mlev > 0) then
		if ForumFType = 0 and ForumStatus > 0 and CatStatus > 0 then
			Response.Write "<a href=""post.asp?method=Topic&amp;FORUM_ID=" & ForumID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>"
		else
			Response.Write "&nbsp;"
		end if
	else
		Response.Write "&nbsp;"
	end if

	if (ForumFType = 0 and ForumACount > 0) and strArchiveState = "1" then
		Response.Write "&nbsp;<a href=""forum.asp?ARCHIVE=true&amp;FORUM_ID=" & ForumID & """>" & getCurrentIcon(strIconFolderArchived,"View Archived posts","") & "</a>"
	end if

	' DEM --> Start of code for Subscription
	if ForumFType = 0 and (strSubscription > 0 and strSubscription < 4) and CatSubscription > 0 and ForumSubscription = 1 and (mlev > 0) and strEmail = 1 then
		if InArray(strForumSubs,ForumID) then
			Response.Write "&nbsp;" & ShowSubLink ("U", ForumCatID, ForumID, 0, "N")
		elseif strBoardSubs <> "Y" and not(InArray(strCatSubs,ForumCatID)) then
			Response.Write "&nbsp;" & ShowSubLink ("S", ForumCatID, ForumID, 0, "N")
		end if
	end if
	' DEM --> End of Code for Subscription
end sub

Sub DoHideCategory(intCatId)
   	HideForumCat = strUniqueID & "HideCat" & intCatId
 	if Request.QueryString(HideForumCat) = "Y" then
 		Response.Cookies(HideForumCat) = "Y"
   		Response.Cookies(HideForumCat).Expires = dateAdd("d", 30, strForumTimeAdjust)
   	else
   		if Request.QueryString(HideForumCat) = "N" then
   			Response.Cookies(HideForumCat) = "N"
   			Response.Cookies(HideForumCat).Expires = dateadd("d", -2, strForumTimeAdjust)
   		end if
   	end if
end sub

Function DoLastPostLink(showicon)
	if ForumLastPostReplyID <> 0 then
		PageLink       = "whichpage=-1"
		AnchorLink     = "&amp;REPLY_ID="
		DoLastPostLink = "<a href=""topic.asp?TOPIC_ID=" & ForumLastPostTopicID & AnchorLink & ForumLastPostReplyID & """>"
		if (showicon = true) then DoLastPostLink = DoLastPostLink & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	elseif ForumLastPostTopicID <> 0 then
		DoLastPostLink = "<a href=""topic.asp?TOPIC_ID=" & ForumLastPostTopicID & """>"
		if (showicon = true) then DoLastPostLink = DoLastPostLink & getCurrentIcon(strIconLastpost,"Jump to Last Post","class=""vam""") & "</a>"
	else
		DoLastPostLink = ""
	end if
end function

function listForumModerators(fForum_ID)
	fForumMods = split(strForumMods,"|")
	for iModerator = 0 to ubound(fForumMods)
		fForumMod  = split(fForumMods(iModerator),",")
		ModForumID = fForumMod(0)
		ModMemID   = fForumMod(1)
		ModMemName = fForumMod(2)
		if cLng(ModForumID) = cLng(fForum_ID) then
			if fMods = "" then
				fMods = profileLink(chkString(ModMemName,"display"),ModMemID)
			else
				fMods = fMods & ", " & profileLink(chkString(ModMemName,"display"),ModMemID)
			end if
		end if
	next
	if fMods = "" then fMods = "&nbsp;"
	listForumModerators = fMods
end function
%>
