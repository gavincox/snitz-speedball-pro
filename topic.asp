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
%>
<!--#INCLUDE FILE="config.asp"-->
<%
if (Request.QueryString("TOPIC_ID") = "" or IsNumeric(Request.QueryString("TOPIC_ID")) = False) and Request.Form("Method_Type") <> "login" and Request.Form("Method_Type") <> "logout" then
	Response.Redirect "default.asp"
	Response.End
else
	Topic_ID = cLng(Request.QueryString("TOPIC_ID"))
end if
Dim ArchiveView, ArchiveLink, CColor
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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%
strLE = "" 'vbNewLine
%>
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="cb/topic_cb.asp" -->
<%
	Response.Write "<div id=""pre-content"">" & strLE & _
		"<div class=""breadcrumbs w50"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
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
		if Forum_Status <> 0 and Cat_Status <> 0 then
			Response.Write getCurrentIcon(strIconFolderOpen,"","class=""vam""")
		else
			Response.Write getCurrentIcon(strIconFolderClosed,"","class=""vam""")
		end if
	end if
	Response.Write "&nbsp;<a href=""forum.asp?" & ArchiveLink & "FORUM_ID=" & Forum_ID & """>" & ChkString(Forum_Subject,"display") & "</a><br>" & strLE
	if ArchiveView = "true" then
		Response.Write getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderArchived,"","class=""vam""") & "&nbsp;"
	elseif Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
		Response.Write getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;"
	else
		Response.Write getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderClosedTopic,"","class=""vam""") & "&nbsp;"
	end if
	if Request.QueryString("SearchTerms") <> "" then
		Response.Write SearchHiLite(ChkString(Topic_Subject,"title"))
	else
		Response.Write ChkString(Topic_Subject,"title")
	end if
	Response.Write "</div>" & strLE & _
		"<!-- /breadcrumbs -->" & strLE & _
		"<div class=""actions w50"">" & strLE
	call PostingOptions()
	Response.Write "</div>" & strLE & _
		"<!-- /actions -->" & strLE & _
		"<div class=""maxpages"">" & strLE
	if maxpages > 1 then
		if mypage > 1 then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&amp;whichpage=" & mypage-1 & SearchLink & """ title=""Goto the Previous page in this Topic"">Previous Page</a>")
		'if mypage > 1 then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage-1 & ";document.PageNum1.submit();"" title=""Goto the Previous page in this Topic"">Previous Page</a>")
		if mypage > 1 and mypage < maxpages then Response.Write(" | ")
		if mypage < maxpages then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&amp;whichpage=" & mypage+1 & SearchLink & """ title=""Goto the Next page in this Topic"">Next Page</a>")
		'if mypage < maxpages then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage+1 & ";document.PageNum1.submit();"" title=""Goto the Next page in this Topic"">Next Page</a>")
	else
		Response.Write "<br style=""font-size: 6px"">" & strLE
	end if
	Response.Write "</div>" & strLE & _
		"<!-- /maxpages -->" & strLE & _
		"</div>" & strLE & _
		"<!-- /pre-content -->" & strLE & strLE & _
		"<table id=""content"">" & strLE  & _
		"<tr>" & strLE & _
		"<th>Author</th>" & strLE & _
		"<th class=""fw"">" & strLE
	if strShowTopicNav = "1" then Call Topic_nav() else Response.Write "Topic"
	Response.Write "</th>" & strLE

	if (AdminAllowed = 1) then
		if maxpages > 1 then
			Response.Write "<th class=""paging"">" & strLE
			Call DropDownPaging(1)
			Response.Write "</th>" & strLE & _
				"<th class=""options"">" & strLE
			Call AdminOptions()
			Response.Write "</th>" & strLE
		else
			Response.Write "<th>&nbsp;</th>" & strLE & _
				"<th class=""options"">" & strLE
			call AdminOptions()
			Response.Write "</th>" & strLE
		end if
	else
		if maxpages > 1 then
			Response.Write "<th class=""paging"">" & strLE
			Call DropDownPaging(1)
			Response.Write "</th>" & strLE & _
				"<th>&nbsp;</th>" & strLE
		else
	        Response.Write "<th>&nbsp;</th><th>&nbsp;</th>" & strLE
	 	end if
	end if
	Response.Write "</tr>" & strLE
	if mypage = 1 then Call GetFirst()
	'## Forum_SQL
	strSql = "UPDATE " & strActivePrefix & "TOPICS "
	strSql = strSql & " SET T_VIEW_COUNT = (T_VIEW_COUNT + 1) "
	strSql = strSql & " WHERE (TOPIC_ID = " & Topic_ID & ")"
	my_conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	if iReplyCount = "" then  '## No replies found in DB
		' Nothing
	else
		intI             = 0
		rM_NAME          = 0
		rM_RECEIVE_EMAIL = 1
		rM_AIM           = 2
		rM_ICQ           = 3
		rM_MSN           = 4
		rM_YAHOO         = 5
		rM_TITLE         = 6
		rMEMBER_ID       = 7
		rM_HOMEPAGE      = 8
		rM_LEVEL         = 9
		rM_POSTS         = 10
		rM_COUNTRY       = 11
		rREPLY_ID        = 12
		rFORUM_ID        = 13
		rR_AUTHOR        = 14
		rTOPIC_ID        = 15
		rR_MESSAGE       = 16
		rR_LAST_EDIT     = 17
		rR_LAST_EDITBY   = 18
		rR_SIG           = 19
		rR_STATUS        = 20
		rR_DATE          = 21
		if CanShowSignature = 1 then rM_SIG = 22
		for iForum = 0 to iReplyCount
			Reply_MemberName         = arrReplyData(rM_NAME, iForum)
			Reply_MemberReceiveEmail = arrReplyData(rM_RECEIVE_EMAIL, iForum)
			Reply_MemberAIM          = arrReplyData(rM_AIM, iForum)
			Reply_MemberICQ          = arrReplyData(rM_ICQ, iForum)
			Reply_MemberMSN          = arrReplyData(rM_MSN, iForum)
			Reply_MemberYAHOO        = arrReplyData(rM_YAHOO, iForum)
			Reply_MemberTitle        = arrReplyData(rM_TITLE, iForum)
			Reply_MemberID           = arrReplyData(rMEMBER_ID, iForum)
			Reply_MemberHomepage     = arrReplyData(rM_HOMEPAGE, iForum)
			Reply_MemberLevel        = arrReplyData(rM_LEVEL, iForum)
			Reply_MemberPosts        = arrReplyData(rM_POSTS, iForum)
			Reply_MemberCountry      = arrReplyData(rM_COUNTRY, iForum)
			Reply_ReplyID            = arrReplyData(rREPLY_ID, iForum)
			Reply_ForumID            = arrReplyData(rFORUM_ID, iForum)
			Reply_Author             = arrReplyData(rR_AUTHOR, iForum)
			Reply_TopicID            = arrReplyData(rTOPIC_ID, iForum)
			Reply_Content            = arrReplyData(rR_MESSAGE, iForum)
			Reply_LastEdit           = arrReplyData(rR_LAST_EDIT, iForum)
			Reply_LastEditBy         = arrReplyData(rR_LAST_EDITBY, iForum)
			Reply_Sig                = arrReplyData(rR_SIG, iForum)
			Reply_Status             = arrReplyData(rR_STATUS, iForum)
			Reply_Date               = arrReplyData(rR_DATE, iForum)
			if CanShowSignature = 1 then Reply_MemberSig = trim(arrReplyData(rM_SIG, iForum))
			if intI = 0 then
				CColor = "fsacc"
			else
				CColor = "ffacc"
			end if
			Response.Write "<tr>" & strLE & _
				"<td class=""poster " & CColor & """ name=""" & Reply_ReplyID & """>" & strLE & _
				"<p><b><span class=""smt"">" & profileLink(ChkString(Reply_MemberName,"display"),Reply_Author) & "</span></b><br>" & strLE
			if strShowRank = 1 or strShowRank = 3 then Response.Write "<span class=""rank"">" & ChkString(getMember_Level(Reply_MemberTitle, Reply_MemberLevel, Reply_MemberPosts),"display") & "</span><br>" & strLE
			if strShowRank = 2 or strShowRank = 3 then Response.Write getStar_Level(Reply_MemberLevel, Reply_MemberPosts) & "<br>" & strLE
		 	Response.Write "</p>" & strLE & _
				"<p class=""rank"">" & strLE
			if strCountry = "1" and trim(Reply_MemberCountry) <> "" then Response.Write Reply_MemberCountry & "<br>" & strLE
			Response.Write Reply_MemberPosts & " Posts</p></td>" & strLE & _
				"<td class=""post " & CColor & """ colspan=""3"">" & strLE
			'"<a> name=""" & Reply_ReplyID & """</a>" & strLE & _
			' DEM --> Start of Code altered for moderation
			if Reply_Status < 2 then
				Response.Write  getCurrentIcon(strIconPosticon,"","class=""vam""") & "&nbsp;<span class=""ffs"">Replied:&nbsp;" & ChkDate(Reply_Date, "&nbsp;-" ,true) & "</span>" & strLE
			elseif Reply_Status = 2 then
				Response.Write  "<span class=""ffs"">NOT MODERATED</span>" & strLE
			elseif Reply_Status = 3 then
				Response.Write  getCurrentIcon(strIconPosticonHold,"","class=""vam""") & "<span class=""ffs"">ON HOLD</span>" & strLE
			end if
			' DEM --> End of Code added for moderation.
			Response.Write "&nbsp;&nbsp;" & profileLink(getCurrentIcon(strIconProfile,"Show Profile","class=""vam"""),Reply_MemberID) & strLE
			if mLev > 2 or Reply_MemberReceiveEmail = "1" then
				if (mlev <> 0) or (mlev = 0 and  strLogonForMail <> "1") then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_mail.asp?id=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconEmail,"Email Poster","class=""vam""") & "</a>" & strLE
				end if
			end if
			if strHomepage = "1" then
				if Reply_MemberHomepage <> " " then
					Response.Write "&nbsp;<a href=""" & Reply_MemberHomepage & """ target=""_blank"">" & getCurrentIcon(strIconHomepage,"Visit " & ChkString(Reply_MemberName,"display") & "'s Homepage","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (AdminAllowed = 1 or Reply_MemberID = MemberID) then
				if (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) or (AdminAllowed = 1) then
					Response.Write "&nbsp;<a href=""post.asp?" & ArchiveLink & "method=Edit&amp;REPLY_ID=" & Reply_ReplyID & "&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconEditTopic,"Edit Reply","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (strAIM = "1") then
				if Trim(Reply_MemberAIM) <> "" then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&amp;ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Reply_MemberName,"display") & " an AOL message","class=""vam""") & "</a>" & strLE
				end if
			end if
			if strICQ = "1" then
				if Trim(Reply_MemberICQ) <> "" then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow6('pop_messengers.asp?mode=ICQ&amp;ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Reply_MemberName,"display") & " an ICQ Message","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (strMSN = "1") then
				if Trim(Reply_MemberMSN) <> "" then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&amp;ID=" & Reply_MemberID & "')"">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Reply_MemberName,"display") & "'s MSN Messenger address","class=""vam""") & "</a>" & strLE
				end if
			end if
			if strYAHOO = "1" then
				if Trim(Reply_MemberYAHOO) <> "" then
					Response.Write "&nbsp;<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Reply_MemberYAHOO, "urlpath") & "&amp;.src=pg"" target=""_blank"">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Reply_MemberName,"display") & " a Yahoo! Message","class=""vam""") & "</a>" & strLE
				end if
			end if
			if ((Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status = 1) or (AdminAllowed = 1 and Topic_Status <= 1)) and ArchiveView = "" then
				Response.Write "&nbsp;<a href=""post.asp?method=ReplyQuote&amp;REPLY_ID=" & Reply_ReplyID & "&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & """>" & getCurrentIcon(strIconReplyTopic,"Reply with Quote","class=""vam""") & "</a>" & strLE
			end if
			if (strIPLogging = "1") then
				if (AdminAllowed = 1) then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_viewip.asp?" & ArchiveLink & "mode=getIP&amp;REPLY_ID=" & Reply_ReplyID & "&amp;FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconIP,"View user's IP address","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (AdminAllowed = 1 or Reply_MemberID = MemberID) then
				if (Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0) or (AdminAllowed = 1) then
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_delete.asp?" & ArchiveLink & "mode=Reply&amp;REPLY_ID=" & Reply_ReplyID & "&amp;TOPIC_ID=" & Topic_ID & "&amp;FORUM_ID=" & Forum_ID & "')"">" & getCurrentIcon(strIconDeleteReply,"Delete Reply","class=""vam""") & "</a>" & strLE
				end if
				' DEM --> Start of Code added for Full Moderation
				if (AdminAllowed = 1 and Reply_Status > 1) then
					ReplyString = "REPLY_ID=" & Reply_ReplyID & "&amp;CAT_ID=" & Cat_ID & "&amp;FORUM_ID=" & Forum_ID & "&amp;TOPIC_ID=" & Topic_ID
					Response.Write "&nbsp;<a href=""JavaScript:openWindow('pop_moderate.asp?" & ReplyString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject this Reply","class=""vam""") & "</a>" & strLE
				end if
				' DEM --> End of Code added for Full Moderation
			end if
			Response.Write "<hr class=""ffs"">" & strLE & _
				"<span class=""smt msg"">"
			if Request.QueryString("SearchTerms") <> "" then
				Response.Write SearchHiLite(formatStr(Reply_Content))
			else
				Response.Write formatStr(Reply_Content)
			end if
			Response.Write "</span>" & strLE & _
				"<!-- /msg -->" & strLE
			if CanShowSignature = 1 and Reply_Sig = 1 and Reply_MemberSig <> "" then
				Response.Write "<hr class=""ffs"">" & strLE & _
					"<span class=""ffc smt"">" & formatStr(Reply_MemberSig) & "</span>" & strLE
			end if
			if strEditedByDate = "1" and Reply_LastEditBy <> "" then
				if Reply_LastEditBy <> Reply_Author then
					Reply_LastEditByName = getMemberName(Reply_LastEditBy)
				else
					Reply_LastEditByName = chkString(Reply_MemberName,"display")
				end if
				Response.Write "<hr class=""ffs"">" & strLE & _
					"<span class=""ffs ffc"">Edited by - " & Reply_LastEditByName & " on " & chkDate(Reply_LastEdit, " " ,true) & "</span>"
			end if
			Response.Write "<a style=""float:right"" href=""#top"">" & getCurrentIcon(strIconGoUp,"Go to Top of Page","") & "</a>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE
			intI  = intI + 1
			if intI = 2 then intI = 0
		next
	end if
	Response.Write "<tr><th>&nbsp;</th>" & strLE & _
		"<th>" & strLE
	if strShowTopicNav = "1" then
		Call Topic_nav()
	else
		Response.Write "Topic"
	end if
	Response.Write "</th>" & strLE
	if (AdminAllowed = 1) then
		if maxpages > 1 then
			Response.Write "<th class=""paging"">" & strLE
			Call DropDownPaging(2)
			Response.Write "</th>" & strLE & _
				"<th class=""options"">" & strLE
			Call AdminOptions()
			Response.Write "</th>" & strLE
		else
			Response.Write "<th>&nbsp;</th>" & strLE & _
				"<th class=""options"">" & strLE
			call AdminOptions()
			Response.Write "</th>" & strLE
		end if
	else
		if maxpages > 1 then
			Response.Write "<th class=""paging"">" & strLE
			Call DropDownPaging(2)
			Response.Write "</th>" & strLE & _
				"<th>&nbsp;</th>" & strLE
		else
	        Response.Write "<th>&nbsp;</th><th>&nbsp;</th>" & strLE
	 	end if
	end if
	Response.Write "</tr>" & strLE & _
		"</table>" & strLE & _
		"<!-- /content -->" & strLE & strLE & _
		"<div id=""post-content"">" & strLE & _
		"<div class=""maxpages"">" & strLE
	if maxpages > 1 then
		if mypage > 1 then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&amp;whichpage=" & mypage-1 & SearchLink & """ title=""Goto the Previous page in this Topic"">Previous Page</a>")
		'if mypage > 1 then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage-1 & ";document.PageNum1.submit();"" title=""Goto the Previous page in this Topic"">Previous Page</a>")
		if mypage > 1 and mypage < maxpages then Response.Write(" | ")
		if mypage < maxpages then Response.Write("<a href=""topic.asp?" & ArchiveLink & "TOPIC_ID=" & Topic_ID & "&amp;whichpage=" & mypage+1 & SearchLink & """ title=""Goto the Next page in this Topic"">Next Page</a>")
		'if mypage < maxpages then Response.Write("<a href=""javascript: onclick=document.PageNum1.whichpage.value=" & mypage+1 & ";document.PageNum1.submit();"" title=""Goto the Next page in this Topic"">Next Page</a>")
	else
		Response.Write "<br style=""font-size: 6px"">" & strLE
	end if
	Response.Write "</div>" & strLE & _
		"<!-- /maxpages -->" & strLE & strLE & _
		"<div class=""actions w50"">" & strLE
	Call PostingOptions()
	Response.Write "</div>" & strLE & _
		"<!-- /actions -->" & strLE & _
		"<div class=""jumpto w50"">" & strLE
%><!--#INCLUDE FILE="inc_jump_to.asp" --><%
	Response.Write "</div>" & strLE & _
		"<!-- /jumpto -->" & strLE
	if strShowQuickReply = "1" and strDBNTUserName <> "" and ((Cat_Status = 1) and (Forum_Status = 1) and (Topic_Status = 1)) and ArchiveView = "" then Call QuickReply()
	Response.Write "</div>" & strLE & _
		"<!-- /post-content -->"  & strLE
	Call WriteFooter
end if
%>