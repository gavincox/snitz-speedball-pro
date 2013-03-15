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
'## Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA
'##
'## Support can be obtained from our support forums at:
'## http://forum.snitz.com
'##
'## Correspondence and Marketing Questions can be sent to:
'## manderson@snitz.com
'##
'#################################################################################

mypage = request("whichpage")
if ((Trim(mypage) = "") or (IsNumeric(mypage) = False)) then mypage = 1
mypage = cLng(mypage)

if Request("SearchTerms") <> "" then
	SearchLink = "&amp;SearchTerms=" & Request("SearchTerms")
else
	SearchLink = ""
end if

if strSignatures = "1" and strDSignatures = "1" then
	if ViewSig(MemberID) <> "0" then
		CanShowSignature = 1
	end if
end if

'## Forum_SQL - Get original topic and check for the Category, Forum or Topic Status and existence
strSql = "SELECT M.M_NAME, M.M_RECEIVE_EMAIL, M.M_AIM, M.M_ICQ, M.M_MSN, M.M_YAHOO" & _
	", M.M_TITLE, M.M_HOMEPAGE, M.MEMBER_ID, M.M_LEVEL, M.M_POSTS, M.M_COUNTRY" & _
	", T.T_DATE, T.T_SUBJECT, T.T_AUTHOR, T.TOPIC_ID, T.T_STATUS, T.T_LAST_EDIT" & _
	", T.T_LAST_EDITBY, T.T_LAST_POST, T.T_SIG, T.T_REPLIES" & _
	", C.CAT_STATUS, C.CAT_ID, C.CAT_NAME, C.CAT_SUBSCRIPTION, C.CAT_MODERATION" & _
	", F.F_STATUS, F.FORUM_ID, F.F_SUBSCRIPTION, F.F_SUBJECT, F.F_MODERATION, T.T_MESSAGE"
if CanShowSignature = 1 then
	strSql = strSql & ", M.M_SIG"
end if
strSql = strSql & " FROM " & strActivePrefix & "TOPICS T, " & strTablePrefix & "FORUM F, " & _
	strTablePrefix & "CATEGORY C, " & strMemberTablePrefix & "MEMBERS M " & _
	" WHERE T.TOPIC_ID = " & Topic_ID & _
	" AND F.FORUM_ID = T.FORUM_ID " & _
	" AND C.CAT_ID = T.CAT_ID " & _
	" AND M.MEMBER_ID = T.T_AUTHOR "

set rsTopic = Server.CreateObject("ADODB.Recordset")
rsTopic.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

if rsTopic.EOF then
	recTopicCount = ""
else
	recTopicCount      = 1
	Member_Name        = rsTopic("M_NAME")
	Member_ReceiveMail = rsTopic("M_RECEIVE_EMAIL")
	Member_AIM         = rsTopic("M_AIM")
	Member_ICQ         = rsTopic("M_ICQ")
	Member_MSN         = rsTopic("M_MSN")
	Member_YAHOO       = rsTopic("M_YAHOO")
	Member_Title       = rsTopic("M_TITLE")
	Member_Homepage    = rsTopic("M_HOMEPAGE")
	TMember_ID         = rsTopic("MEMBER_ID")
	Member_Level       = rsTopic("M_LEVEL")
	Member_Posts       = rsTopic("M_POSTS")
	Member_Country     = rsTopic("M_COUNTRY")
	Topic_Date         = rsTopic("T_DATE")
	Topic_Subject      = rsTopic("T_SUBJECT")
	Topic_Author       = rsTopic("T_AUTHOR")
	TopicID            = rsTopic("TOPIC_ID")
	Topic_Status       = rsTopic("T_STATUS")
	Topic_LastEdit     = rsTopic("T_LAST_EDIT")
	Topic_LastEditby   = rsTopic("T_LAST_EDITBY")
	Topic_LastPost     = rsTopic("T_LAST_POST")
	Topic_Sig          = rsTopic("T_SIG")
	Topic_Replies      = rsTopic("T_REPLIES")
	Cat_Status         = rsTopic("CAT_STATUS")
	Cat_ID             = rsTopic("CAT_ID")
	Cat_Name           = rsTopic("CAT_NAME")
	Cat_Subscription   = rsTopic("CAT_SUBSCRIPTION")
	Cat_Moderation     = rsTopic("CAT_MODERATION")
	Forum_Status       = rsTopic("F_STATUS")
	Forum_ID           = rsTopic("FORUM_ID")
	Forum_Subject      = rsTopic("F_SUBJECT")
	Forum_Subscription = rsTopic("F_SUBSCRIPTION")
	Forum_Moderation   = rsTopic("F_MODERATION")
	Topic_Message      = rsTopic("T_MESSAGE")
	if CanShowSignature = 1 then Topic_MemberSig = trim(rsTopic("M_SIG"))
end if

rsTopic.close
set rsTopic = nothing

if recTopicCount = "" then
	if ArchiveView <> "true" then
		Response.Redirect("topic.asp?ARCHIVE=true&amp;" & ChkString(Request.QueryString,"sqlstring"))
	else
		Response.Redirect("default.asp")
	end if
end if

if mLev = 4 then
	AdminAllowed = 1
	ForumChkSkipAllowed = 1
elseif mLev = 3 then
	if chkForumModerator(Forum_ID, chkString(strDBNTUserName,"decode")) = "1" then
		AdminAllowed = 1
		ForumChkSkipAllowed = 1
	else
		if lcase(strNoCookies) = "1" then
			AdminAllowed = 1
			ForumChkSkipAllowed = 0
		else
			AdminAllowed = 0
			ForumChkSkipAllowed = 0
		end if
	end if
elseif lcase(strNoCookies) = "1" then
 	AdminAllowed = 1
	ForumChkSkipAllowed = 0
else
 	AdminAllowed = 0
	ForumChkSkipAllowed = 0
end if

if strPrivateForums = "1" and (Request.Form("Method_Type") <> "login") and (Request.Form("Method_Type") <> "logout") and ForumChkSkipAllowed = 0 then
	result = ChkForumAccess(Forum_ID, MemberID, true)
end if

if strModeration > 0 and Cat_Moderation > 0 and Forum_Moderation > 0 and AdminAllowed = 0 then
	Moderation = "Y"
else
	Moderation = "N"
end if

if mypage = -1 then
	strSql = "SELECT REPLY_ID FROM " & strActivePrefix & "REPLY WHERE TOPIC_ID = " & Topic_ID & " "
	if AdminAllowed = 0 then
		strSql = strSql & " AND (R_STATUS < "
		if Moderation = "Y" then
			strSql = strSql & "2 "
		else
			strSql = strSql & "3 "
		end if
		strSql = strSql & "OR R_AUTHOR = " & MemberID & ") "
	end if
	strSql = strSql & "ORDER BY R_DATE ASC "

	set rsReplies = Server.CreateObject("ADODB.Recordset")
	if strDBType = "mysql" then
		rsReplies.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	else
		rsReplies.open strSql, my_Conn, adOpenStatic, adLockReadOnly, adCmdText
	end if

	if not rsReplies.EOF then
		arrReplyData = rsReplies.GetRows(adGetRowsRest)
		iReplyCount = UBound(arrReplyData, 2)

		if Request.Querystring("REPLY_ID") <> "" and IsNumeric(Request.Querystring("REPLY_ID")) then
			LastPostReplyID = cLng(Request.Querystring("REPLY_ID"))
			for iReply = 0 to iReplyCount
				intReplyID = arrReplyData(0, iReply)
				if LastPostReplyID = intReplyID then
					intPageNumber = ((iReply+1)/strPageSize)
					exit for
				end if
			next
		else
			LastPostReplyID = cLng(arrReplyData(0, iReplyCount))
			intPageNumber = ((iReplyCount+1)/strPageSize)
		end if
		if intPageNumber > cLng(intPageNumber) then
			intPageNumber = cLng(intPageNumber) + 1
		end if
		strwhichpage = "whichpage=" & intPageNumber & "&amp;"
	else
		strwhichpage = ""
	end if

	rsReplies.close
	set rsReplies = nothing
	my_Conn.close
	set my_Conn   = nothing

	Response.Redirect "topic.asp?" & ArchiveLink & strwhichpage & "TOPIC_ID=" & Topic_ID & SearchLink & "#" & LastPostReplyID
	Response.End
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
end If

Response.Write "<script type=""text/javascript"">" & _
	"function ChangePage(fnum){" & _
	"if (fnum == 1) {" & _
	"document.PageNum1.submit();" & _
	"}" & _
	"else {" & _
	"document.PageNum2.submit();" & _
	"}" & _
	"}" & _
	"</script>"

if (Moderation = "Y" and Topic_Status > 1 and Topic_Author <> MemberID) then
	Response.write  "<div class=""lmessage""><p>Viewing of this Topic is not permitted until it has been moderated.</p><p>Please try again later</p>" & strLE & _
		"<a href=""JavaScript:history.go(-1)"">Go Back</a>" & strLE & _
		"</div>" & strLE & _
		"<!-- /lmessage --><br>" & strLE
	WriteFooter
	Response.end
else
	'## Forum_SQL
	strSql = "SELECT M.M_NAME, M.M_RECEIVE_EMAIL, M.M_AIM, M.M_ICQ, M.M_MSN, M.M_YAHOO"
	strSql = strSql & ", M.M_TITLE, M.MEMBER_ID, M.M_HOMEPAGE, M.M_LEVEL, M.M_POSTS, M.M_COUNTRY"
	strSql = strSql & ", R.REPLY_ID, R.FORUM_ID, R.R_AUTHOR, R.TOPIC_ID, R.R_MESSAGE, R.R_LAST_EDIT"
	strSql = strSql & ", R.R_LAST_EDITBY, R.R_SIG, R.R_STATUS, R.R_DATE"
	if CanShowSignature = 1 then
		strSql = strSql & ", M.M_SIG"
	end if
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS M, " & strActivePrefix & "REPLY R "
	strSql3 = " WHERE M.MEMBER_ID = R.R_AUTHOR "
	strSql3 = strSql3 & " AND R.TOPIC_ID = " & Topic_ID & " "
		' DEM --> if not a Moderator, all unapproved posts should not be viewed.
		if AdminAllowed = 0 then
			strSql3 = strSql3 & " AND (R.R_STATUS < "
			if Moderation = "Y" then
				' Ignore unapproved/rejected posts
				strSql3 = strSql3 & "2"
			else
				' Ignore any previously rejected topic
				strSql3 = strSql3 & "3"
			end if
		strSql3 = strSql3 & " OR R.R_AUTHOR = " & MemberID & ")"
	end if
	strSql4 = " ORDER BY R.R_DATE ASC"

	if strDBType = "mysql" then 'MySql specific code
		if mypage > 1 then
 			intOffset = cLng((mypage-1) * strPageSize)
			strSql5 = " LIMIT " & intOffset & ", " & strPageSize & " "
		end if

		'## Forum_SQL - Get the total pagecount
		strSql1 = "SELECT COUNT(R.TOPIC_ID) AS REPLYCOUNT "

		set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
		iPageTotal = rsCount(0).value
		rsCount.close
		set rsCount = nothing

		if iPageTotal > 0 then
			maxpages = (iPageTotal  \ strPageSize )
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
			iPageTotal = 0
			maxpages = 0
		end if

		if iPageTotal > 0 then
			set rsReplies = Server.CreateObject("ADODB.Recordset")
			rsReplies.Open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
				arrReplyData = rsReplies.GetRows(intGetRows)
				iReplyCount = UBound(arrReplyData, 2)
			rsReplies.Close
			set rsReplies = nothing
		else
			iReplyCount = ""
		end if

	else 'end MySql specific code

		set rsReplies = Server.CreateObject("ADODB.Recordset")
		rsReplies.cachesize = strPageSize
		rsReplies.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic, adLockReadOnly, adCmdText

			if not (rsReplies.EOF or rsReplies.BOF) then
				rsReplies.pagesize = strPageSize
				rsReplies.absolutepage = mypage '**
				maxpages = cLng(rsReplies.pagecount)
				if maxpages >= mypage then
					arrReplyData = rsReplies.GetRows(strPageSize)
					iReplyCount = UBound(arrReplyData, 2)
				else
					iReplyCount = ""
				end if
			else  '## No replies found in DB
				iReplyCount = ""
			end if

		rsReplies.Close
		set rsReplies = nothing
	end if

sub GetFirst()
	Response.Write "<tr>" & strLE & _
		"<td class=""poster ffcc"">" & strLE & _
		"<p><b><span class=""smt"">" & profileLink(ChkString(Member_Name,"display"),TMember_ID) & "</span></b><br>" & strLE
	if strShowRank = 1 or strShowRank = 3 then Response.Write "<span class=""rank"">" & ChkString(getMember_Level(Member_Title, Member_Level, Member_Posts),"display") & "</span><br>" & strLE
	if strShowRank = 2 or strShowRank = 3 then Response.Write getStar_Level(Member_Level, Member_Posts) & "<br>" & strLE
 	Response.Write "</p>" & strLE & _
		"<p class=""rank"">" & strLE
	if strCountry = "1" and trim(Member_Country) <> "" then Response.Write Member_Country & "<br>" & strLE
	Response.Write Member_Posts & " Posts</p></td>" & strLE
	Response.Write "<td class=""post ffcc"" colspan=""3"">"
	if Topic_Status < 2 then
		Response.Write   getCurrentIcon(strIconPosticon,"","class=""vam""") & "&nbsp;<span class=""ffs"">Posted:&nbsp;" & ChkDate(Topic_Date, "&nbsp;-" ,true) & "</span>" & strLE
	elseif Topic_Status = 2 then
		Response.Write  "<span class=""ffs"">NOT MODERATED</span>" & strLE
	elseif Topic_Status = 3 then
		Response.Write  getCurrentIcon(strIconPosticonHold,"","valign=""vam""") & "<span class=""ffs"">ON HOLD</span>" & strLE
	end if
	Response.Write "&nbsp;&nbsp;" & profileLink(getCurrentIcon(strIconProfile,"Show Profile","class=""vam"""),TMember_ID) & strLE
	if mLev > 2 or Member_ReceiveMail = "1" then
		if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_mail.asp?id=" & TMember_ID & "')"">" & getCurrentIcon(strIconEmail,"Email Poster","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (strHomepage = "1") then
		if Member_Homepage <> " " then
			Response.Write "&nbsp;<a href=""" & Member_Homepage & """ target=""_blank"">" & getCurrentIcon(strIconHomepage,"Visit " & ChkString(Member_Name,"display") & "'s Homepage","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (AdminAllowed = 1 or TMember_ID = MemberID) then
		if ((Cat_Status <> 0) and (Forum_Status <> 0) and (Topic_Status <> 0)) or (AdminAllowed = 1) then
			Response.Write "&nbsp;<a href=""post.asp?" & ArchiveLink & "method=EditTopic&amp;REPLY_ID=" & Topic_ID & "&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconEditTopic,"Edit Topic","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (strAIM = "1") then
		if Trim(Member_AIM) <> "" then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&amp;ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Member_Name,"display") & " an AOL message","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (strICQ = "1") then
		if Trim(Member_ICQ) <> "" then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow6('pop_messengers.asp?mode=ICQ&amp;ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Member_Name,"display") & " an ICQ Message","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (strMSN = "1") then
		if Trim(Member_MSN) <> "" then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&amp;ID=" & TMember_ID & "')"">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Member_Name,"display") & "'s MSN Messenger address","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (strYAHOO = "1") then
		if Trim(Member_YAHOO) <> "" then
			Response.Write "&nbsp;<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Member_YAHOO, "urlpath") & "&amp;.src=pg"" target=""_blank"">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Member_Name,"display") & " a Yahoo! Message","class=""vam""") & "</a>" & strLE
		end if
	end if
	if ((Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status = 1) or (AdminAllowed = 1 and Topic_Status <= 1) and ArchiveView = "" ) then
		Response.Write "&nbsp;<a href=""post.asp?method=TopicQuote&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply with Quote","class=""vam""") & "</a>" & strLE
	end if
	if (strIPLogging = "1") then
		if (AdminAllowed = 1) then
			Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_viewip.asp?" & ArchiveLink & "mode=getIP&amp;TOPIC_ID=" & TopicID & "&amp;FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconIP,"View user's IP address","class=""vam""") & "</a>" & strLE
		end if
	end if
	if (AdminAllowed = 1) or (TMember_ID = MemberID and Topic_Replies < 1) then
		Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconDeleteReply,"Delete Topic","class=""vam""") & "</a>" & strLE
	end if
	' DEM --> Start of Code added for Full Moderation
	if (AdminAllowed = 1 and Topic_Status > 1) then
		TopicString = "TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID
		Response.Write  "&nbsp;<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Topic","class=""vam""") & "</a>" & strLE
	End if
	Response.Write "<hr class=""ffs"">" & strLE & _

		"<span class=""smt msg"">"
	if Request.QueryString("SearchTerms") <> "" then
		Response.Write SearchHiLite(formatStr(Topic_Message))
	else
		Response.Write formatStr(Topic_Message)
	end if
	Response.Write "</span>" & strLE & _
	"<!-- .msg -->" & strLE

	if CanShowSignature = 1 and Topic_Sig = 1 and Topic_MemberSig <> "" then
		Response.Write "<hr class=""ffs"">" & strLE & _
		"<span class=""ffc smt"">" & formatStr(Topic_MemberSig) & "</span>" & strLE
	end if
	if strEditedByDate = "1" and Topic_LastEditBy <> "" then
		if Topic_LastEditBy <> Topic_Author then
			Topic_LastEditByName = getMemberName(Topic_LastEditBy)
		else
			Topic_LastEditByName = chkString(Member_Name,"display")
		end if
		Response.Write "<hr class=""ffs"">" & strLE & _
		"<span class=""ffs ffc"">Edited by - " & Topic_LastEditByName & " on " & chkDate(Topic_LastEdit, " " ,true) & "</span>"
	end if
	Response.Write "<a style=""float:right"" href=""#top"">" & getCurrentIcon(strIconGoUp,"Go to Top of Page","") & "</a>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE
End Sub


sub PostingOptions()
	if (mlev = 4 or mlev = 3 or mlev = 2 or mlev = 1) _
	or (lcase(strNoCookies) = "1") or (strDBNTUserName = "") then
		if ((Cat_Status = 1) and (Forum_Status = 1)) then
			Response.Write "<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>New Topic</a></span>" & strLE
		else
			if (AdminAllowed = 1) then
				Response.Write "<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderLocked,"New Topic","class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>New Topic</a></span>" & strLE
			else
				Response.Write getCurrentIcon(strIconFolderLocked,"Forum Locked","class=""vam""") & "<span class=""label"">&nbsp;&nbsp;Forum Locked</span>" & strLE
			end if
		end if
		if ((Cat_Status = 1) and (Forum_Status = 1) and (Topic_Status = 1)) and ArchiveView = "" then
			Response.Write "<a href=""post.asp?method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic", "class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""post.asp?method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>Reply to Topic</a><br></span>" & strLE
		else
			if ((AdminAllowed = 1 and Topic_Status <= 1) and ArchiveView = "")  then
				Response.Write "<a href=""post.asp?method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>"
				' DEM --> Added if statement to show normal icon for unmoderated posts.
				if Topic_Status = 1 and Cat_Status <> 0 and Forum_Status <> 0 then
					Response.Write getCurrentIcon(strIconReplyTopic,"Reply to Topic", "class=""vam""") & "</a>&nbsp;"
				else
					Response.Write getCurrentIcon(strIconClosedTopic,"Reply to Topic", "class=""vam""") & "</a>&nbsp;"
				end if
				Response.Write "<span class=""label""><a href=""post.asp?" & ArchiveLink & "method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>Reply to Topic</a><br></span>" & strLE
			else
				if Topic_Status = 0 then
					Response.Write getCurrentIcon(strIconClosedTopic,"Topic Locked", "class=""vam""") & "<span class=""label"">&nbsp;Topic Locked<br></span>" & strLE
				end if
			end if
		end if
		if lcase(strEmail) = "1" and Topic_Status < 2 then
			if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 and mLev > 0 then
				if strSubscription > 0 and Cat_Subscription > 0 and Forum_Subscription > 0 then
					if InArray(strTopicSubs, Topic_ID) then
						Response.Write "<br>" & ShowSubLink ("U", Cat_ID, Forum_ID, Topic_ID, "Y") & strLE
					elseif strBoardSubs <> "Y" and not(InArray(strForumSubs,Forum_ID) or InArray(strCatSubs,Cat_ID)) then
						Response.Write "<br>" & ShowSubLink ("S", Cat_ID, Forum_ID, Topic_ID, "Y") & strLE
					end if
				end if
			end if
			if ((mlev <> 0) or (mlev = 0 and strLogonForMail <> "1")) and lcase(strShowSendToFriend) = "1" then
				Response.Write "<a href=""JavaScript:openWindow('pop_send_to_friend.asp?url=" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "')"">" & getCurrentIcon(strIconSendTopic,"Send Topic to a Friend","class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""JavaScript:openWindow('pop_send_to_friend.asp?url=" & strForumURL & "topic.asp?TOPIC_ID=" & Topic_ID & "')"">Send Topic to a Friend</a><br></span>" & strLE
			end if
		end if
		if lcase(strShowPrinterFriendly) = "1" and Topic_Status < 2 then
			Response.Write "<a href=""JavaScript:openWindow5('pop_printer_friendly.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "')"">" & getCurrentIcon(strIconPrint,"Printer Friendly","class=""vam""") & "</a><span class=""label"">&nbsp;<a href=""JavaScript:openWindow5('pop_printer_friendly.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "')"">Printer Friendly</a></span>" & strLE
		end if
	end if
end sub

sub AdminOptions()
	if (AdminAllowed = 1) or (lcase(strNoCookies) = "1") then
		if (Cat_Status = 0) then
			if (mlev = 4) then
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Category&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Category","") & "</a>" & strLE
			else
				Response.Write "" & getCurrentIcon(strIconFolderUnlocked,"Category Locked","") & strLE
			end if
		else
			if (Forum_Status = 0) then
				Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Forum&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Forum","") & "</a>" & strLE
			else
				if (Topic_Status <> 0) then
					Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderLocked,"Lock Topic","") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderUnlocked,"Un-Lock Topic","") & "</a>" & strLE
				end if
			end if
		end if
		if ((Cat_Status <> 0) and (Forum_Status <> 0) and (Topic_Status <> 0)) or (AdminAllowed = 1) then
			Response.Write "<a href=""post.asp?" & ArchiveLink & "method=EditTopic&amp;REPLY_ID=" & Topic_ID & "&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderPencil,"Edit Topic","") & "</a>" & strLE
		end if
		Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Topic&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "')"">" & getCurrentIcon(strIconFolderDelete,"Delete Topic","") & "</a>" & strLE & _
			"<a href=""post.asp?method=Topic&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconFolderNewTopic,"New Topic","") & "</a>" & strLE
		if Topic_Status <= 1 and ArchiveView = "" then
			Response.Write "<a href=""post.asp?method=Reply&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply to Topic","") & "</a>" & strLE
		end if
	end if
	' DEM --> Start of Code added for Full Moderation
	if (AdminAllowed = 1 and CheckForUnModeratedPosts("TOPIC", Cat_ID, Forum_ID, Topic_ID) > 0) then
		TopicString = "TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;CAT_ID=" & Cat_ID & "&amp;REPLY_ID=X"
		Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject all posts for this Topic","") & "</a>" & strLE
	end if
	' DEM --> End of Code added for Full Moderation
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write "<form name=""PageNum" & fnum & """ action=""topic.asp"">" & strLE
		if Archiveview = "true" then Response.Write "<input type=""hidden"" name=""ARCHIVE"" value=""" & ArchiveView & """>" & strLE
		Response.Write "<input type=""hidden"" name=""TOPIC_ID"" value=""" & Request("TOPIC_ID") & """>" & strLE
		Response.Write "Page:&nbsp;<select name=""whichpage"" size=""1"" onChange=""ChangePage(" & fnum & ");"">" & strLE
		for counter = 1 to maxpages
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		Response.Write "</select> of " & maxpages & strLE
		if Request.QueryString("SearchTerms") <> "" then
			Response.Write "<input type=""hidden"" name=""SearchTerms"" value=""" & Server.HTMLEncode(Request.QueryString("SearchTerms")) & """>" & strLE
		end if
		Response.Write "</form>" & strLE
	end if
	top = "0"
end sub

Sub Topic_nav()
	if prevTopic = "" then
		strSQL = "SELECT T_SUBJECT, TOPIC_ID "
		strSql = strSql & "FROM " & strActivePrefix & "TOPICS "
		strSql = strSql & "WHERE T_LAST_POST > '" & Topic_LastPost
		strSql = strSql & "' AND FORUM_ID = " & Forum_ID
		strSql = strSql & " AND T_STATUS < 2"  ' Ignore unapproved/held posts
		strSql = strSql & " ORDER BY T_LAST_POST;"

		set rsPrevTopic = my_conn.Execute(TopSQL(strSql,1))

		if rsPrevTopic.EOF then
			prevTopic = getCurrentIcon(strIconBlank,"","class=""vam""")
		else
			prevTopic = "&nbsp;<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & rsPrevTopic("TOPIC_ID") & """>" & getCurrentIcon(strIconGoLeft,"Previous Topic","class=""vam""") & "</a>&nbsp;"
		end if

		rsPrevTopic.close
		set rsPrevTopic = nothing
	else
		prevTopic = prevTopic
	end if

	if NextTopic = "" then
		strSQL = "SELECT T_SUBJECT, TOPIC_ID "
		strSql = strSql & "FROM " & strActivePrefix & "TOPICS "
		strSql = strSql & "WHERE T_LAST_POST < '" & Topic_LastPost
		strSql = strSql & "' AND FORUM_ID = " & Forum_ID
		strSql = strSql & " AND T_STATUS < 2"  ' Ignore unapproved/held posts
		strSql = strSql & " ORDER BY T_LAST_POST DESC;"

		set rsNextTopic = my_conn.Execute(TopSQL(strSql,1))

		if rsNextTopic.EOF then
			nextTopic = getCurrentIcon(strIconBlank,"","class=""vam""")
		else
			nextTopic = "&nbsp;<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & rsNextTopic("TOPIC_ID") & """>" & getCurrentIcon(strIconGoRight,"Next Topic","class=""vam""") & "</a>&nbsp;"
		end if

		rsNextTopic.close
		set rsNextTopic = nothing
	else
		nextTopic = nextTopic
	end if

	Response.Write prevTopic & "&nbsp;Topic&nbsp;" & nextTopic
end sub

function SearchHiLite(fStrMessage)
	'function derived from HiLiTeR by 2eNetWorX
	fArr = split(replace(Request.QueryString("SearchTerms"),";",""), ",")
	strBuffer = ""
	for iPos = 1 to len(fStrMessage)
		bChange = False
		'Looks for html tags
		if mid(fStrMessage, iPos, 1) = "<" then
			bInHTML = True
		end if
		'Looks for End of html tags
		if bInHTML = True then
			if mid(fStrMessage, iPos, 1) = ">" then
				bInHTML = False
			end if
		end if
		if bInHTML <> True then
			for i = 0 to UBound(fArr)
				if fArr(i) <> "" then
					if lcase(mid(fStrMessage, iPos, len(fArr(i)))) = lcase(fArr(i)) then
						bChange   = True
						strBuffer = strBuffer & "<span class=""spnSearchHighlight"" id=""hilite"">" & _
						mid(fStrMessage, iPos, len(fArr(i))) & "</span id=""hilite"">"
						iPos = iPos + len(fArr(i)) - 1
					end if
				end if
			next
		end if
		if Not bChange then
			strBuffer = strBuffer & mid(fStrMessage, iPos, 1)
		end if
	next
	SearchHiLite = strBuffer
end function

Sub QuickReply()
	intSigDefault = getSigDefault(MemberID)
	Response.Write "<script type=""text/javascript"" src=""inc_code.js""></script>" & strLE & _
		"<div style=""clear:both""><br style=""font-size: 6px""></div>" & strLE & _
		"<form name=""PostTopic"" method=""post"" action=""post_info.asp"" onSubmit=""return validate();"">" & strLE & _
		"<input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & strLE & _
		"<input name=""Method_Type"" type=""hidden"" value=""Reply"">" & strLE & _
		"<input name=""TOPIC_ID"" type=""hidden"" value=""" & Topic_ID & """>" & strLE & _
		"<input name=""FORUM_ID"" type=""hidden"" value=""" & Forum_ID & """> " & strLE & _
		"<input name=""CAT_ID"" type=""hidden"" value=""" & Cat_ID & """>" & strLE & _
		"<input name=""Refer"" type=""hidden"" value=""" & request.servervariables("SCRIPT_NAME") & "?" & chkString(Request.QueryString,"refer") & """>" & strLE & _
		"<input name=""UserName"" type=""hidden"" value=""" & strDBNTUserName & """>" & strLE & _
		"<input name=""Password"" type=""hidden"" value=""" & Request.Cookies(strUniqueID & "User")("Pword") & """>" & strLE & _
		"<table id=""quickr"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2"">Quick Reply</th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc vat""><span class=""smt r"" style=""display:block""><b>Message:&nbsp;</b><br>" & strLE & _
		"<br>" & strLE & _
		"<span class=""dff ffs ffc nw l"">" & strLE
	if strAllowHTML = "1" then
		Response.Write "* HTML is ON<br>" & strLE
	else
		Response.Write "* HTML is OFF<br>" & strLE
	end if
	if strAllowForumCode = "1" then
		Response.Write "* <a href=""JavaScript:openWindow6('pop_forum_code.asp')"">Forum Code</a> is ON<br>" & strLE
	else
		Response.Write "* Forum Code is OFF<br>" & strLE
	end if
	if strSignatures = "1" then
		Response.Write "<br><input name=""Sig"" id=""Sig"" type=""checkbox"" value=""yes""" & chkCheckbox(intSigDefault,1,true) & "><label for=""Sig"">Include Signature</label><br>" & strLE
	end if
	Response.Write "</span></td>" & strLE & _
		"<td class=""fcc fw""><textarea name=""Message"" cols=""50"" rows=""6"" wrap=""virtual"" class=""w95""></textarea><br></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc c nw"" colspan=""2""><input name=""Submit"" type=""submit"" value=""Submit Reply"">&nbsp;<input name=""Preview"" type=""button"" value=""Preview Reply"" onclick=""OpenPreview()""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE & _
		"<br>" & strLE
end sub
%>
