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
Dim HeldFound, UnApprovedFound, UnModeratedPosts, UnModeratedFPosts
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_chknew.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%
if mLev < 3 then Response.Redirect("default.asp")

Dim ArchiveView
if request("ARCHIVE") = "true" then
	strActivePrefix = strArchiveTablePrefix
	ArchiveView     = "true"
	ArchiveLink     = "ARCHIVE=true&"
else
	strActivePrefix = strTablePrefix
	ArchiveView     = ""
	ArchiveLink     = ""
end if

Response.Write "<script type=""text/javascript"">" & strLE & _
	"function ChangePage(fnum){" & strLE & _
	"if (fnum == 1) {" & strLE & _
	"document.PageNum1.submit();" & strLE & _
	"}" & strLE & _
	"else {" & strLE & _
	"document.PageNum2.submit();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"</script>" & strLE

' DEM --> Added code for topic moderation
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

'## Forum_SQL - Find all records with the search criteria in them
strSql = "SELECT DISTINCT C.CAT_STATUS, C.CAT_NAME, C.CAT_ORDER"
strSql = strSql & ", F.F_ORDER, F.FORUM_ID, F.F_SUBJECT, F.CAT_ID, F.F_PRIVATEFORUMS"
strSql = strSql & ", F.F_PASSWORD_NEW, F.F_STATUS"
strSql = strSql & ", T.TOPIC_ID, T.T_AUTHOR, T.T_SUBJECT, T.T_STATUS, T.T_LAST_POST"
strSql = strSql & ", T.T_LAST_POST_AUTHOR, T.T_REPLIES, T.T_UREPLIES, T.T_VIEW_COUNT"
strSql = strSql & ", M.MEMBER_ID, M.M_NAME, MEMBERS_1.M_NAME AS LAST_POST_AUTHOR_NAME "

strSql2 = " FROM ((((" & strTablePrefix & "FORUM F LEFT JOIN " & strActivePrefix & "TOPICS T"
strSql2 = strSql2 & " ON F.FORUM_ID = T.FORUM_ID) LEFT JOIN " & strActivePrefix & "REPLY R"
strSql2 = strSql2 & " ON T.TOPIC_ID = R.TOPIC_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS M"
strSql2 = strSql2 & " ON T.T_AUTHOR = M.MEMBER_ID) LEFT JOIN " & strTablePrefix & "CATEGORY C"
strSql2 = strSql2 & " ON T.CAT_ID = C.CAT_ID) LEFT JOIN " & strMemberTablePrefix & "MEMBERS MEMBERS_1"
strSql2 = strSql2 & " ON T.T_LAST_POST_AUTHOR = MEMBERS_1.MEMBER_ID"

strSql3 = " WHERE (T.T_STATUS > 1 OR R.R_STATUS > 1)"
if mlev = 3 and ModOfForums <> "" then strSql3 = strSql3 & " AND T.FORUM_ID IN (" & ModOfForums & ") "

strSql4 = " ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT, T.T_LAST_POST DESC"

mypage = request("whichpage")
if ((Trim(mypage) = "") or (IsNumeric(mypage) = False)) then mypage = 1
mypage = cLng(mypage)

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		intOffset = cLng((mypage-1) * strPageSize)
		strSql5 = strSql5 & " LIMIT " & intOffset & ", " & strPageSize & " "
	end if

	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(DISTINCT T.TOPIC_ID) AS PAGECOUNT "

	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing

	if iPageTotal > 0 then
		inttotaltopics = iPageTotal
		maxpages = (iPageTotal \ strPageSize )
		if iPageTotal mod strPageSize <> 0 then maxpages = maxpages + 1
		if iPageTotal < (strPageSize + 1) then
			intGetRows = iPageTotal
		elseif (mypage * strPageSize) > iPageTotal then
			intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
		else
			intGetRows = strPageSize
		end if
	else
		iPageTotal = 0
		inttotaltopics = iPageTotal
		maxpages = 0
	end if

	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			arrTopicData = rs.GetRows(intGetRows)
			iTopicCount  = UBound(arrTopicData, 2)
		rs.close
		set rs = nothing
	else
		iTopicCount = ""
	end if

else 'end MySql specific code

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		if not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize     = strPageSize
			inttotaltopics  = cLng(rs.recordcount)
			rs.absolutepage = mypage '**
			maxpages        = cLng(rs.pagecount)
			arrTopicData    = rs.GetRows(strPageSize)
			iTopicCount     = UBound(arrTopicData, 2)
		else
			iTopicCount    = ""
			inttotaltopics = 0
		end if
	rs.Close
	set rs = nothing
end if

if strModeration = "1" and mLev > 2 then UnModeratedPosts = CheckForUnmoderatedPosts("BOARD", 0, 0, 0)

Response.Write "<table class=""tc"" width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw l"" width=""33%""><span class=""dff dfs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & " <a href=""moderate.asp"">Unmoderated Posts</a></span></td>" & strLE & _
	"</tr>" & strLE
if maxpages > 1 then
	Response.Write "<tr class=""r"">" & strLE
	Call DropDownPaging(1)
	Response.Write "</tr>" & strLE
end if
Response.Write "</table>" & strLE

Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">&nbsp;</span></b></td>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Topic</span></b></td>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Author</span></b></td>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Replies</span></b></td>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Read</span></b></td>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Last Post</span></b></td>" & strLE
if (mlev > 0) or (lcase(strNoCookies) = "1") then
	Response.Write "<td class=""hcc c""><b><span class=""dff dfs hfc"">"
	if (mLev = 4) then
		if UnModeratedPosts > 0 then
			Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts","class=""vam""") & "</a>"
		else
			Response.Write("&nbsp;")
		end if
	else
		Response.Write("&nbsp;")
	end if
	Response.Write "</span></b></td>" & strLE
end if
Response.Write "</tr>" & strLE
if iTopicCount = "" then '## No Search Results
	Response.Write "<tr>" & strLE & _
		"<td class=""fcc"" colspan=""6""><span class=""dff dfs""><b>No Matches Found</b></span></td>" & strLE
	if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then Response.Write "<td class=""fcc c""><b><span class=""dff dfs hfc"">&nbsp;</span></b></td>" & strLE
	Response.Write "</tr>" & strLE
else
	cCAT_STATUS            = 0
	cCAT_NAME              = 1
	fFORUM_ID              = 4
	fF_SUBJECT             = 5
	fCAT_ID                = 6
	fF_PRIVATEFORUMS       = 7
	fF_PASSWORD_NEW        = 8
	fF_STATUS              = 9
	tTOPIC_ID              = 10
	tT_AUTHOR              = 11
	tT_SUBJECT             = 12
	tT_STATUS              = 13
	tT_LAST_POST           = 14
	tT_LAST_POST_AUTHOR    = 15
	tT_REPLIES             = 16
	tT_UREPLIES            = 17
	tT_VIEW_COUNT          = 18
	mMEMBER_ID             = 19
	mM_NAME                = 20
	tLAST_POST_AUTHOR_NAME = 21

	currForum              = 0
	currTopic              = 0
	dim Cat_Status
	dim Forum_Status
	dim mdisplayed
	mdisplayed             = 0
	rec                    = 1

	for iTopic = 0 to iTopicCount
		if (rec = strPageSize + 1) then exit for

		Cat_Status               = arrTopicData(cCAT_STATUS, iTopic)
		Cat_Name                 = arrTopicData(cCAT_NAME, iTopic)
		Forum_ID                 = arrTopicData(fFORUM_ID, iTopic)
		Forum_Subject            = arrTopicData(fF_SUBJECT, iTopic)
		Forum_Cat_ID             = arrTopicData(fCAT_ID, iTopic)
		Forum_PrivateForums      = arrTopicData(fF_PRIVATEFORUMS, iTopic)
		Forum_FPasswordNew       = arrTopicData(fF_PASSWORD_NEW, iTopic)
		Forum_Status             = arrTopicData(fF_STATUS, iTopic)
		Topic_ID                 = arrTopicData(tTOPIC_ID, iTopic)
		Topic_Author             = arrTopicData(tT_AUTHOR, iTopic)
		Topic_Subject            = arrTopicData(tT_SUBJECT, iTopic)
		Topic_Status             = arrTopicData(tT_STATUS, iTopic)
		Topic_LastPost           = arrTopicData(tT_LAST_POST, iTopic)
		Topic_LastPostAuthor     = arrTopicData(tT_LAST_POST_AUTHOR, iTopic)
		Topic_Replies            = arrTopicData(tT_REPLIES, iTopic)
		Topic_UReplies           = arrTopicData(tT_UREPLIES, iTopic)
		Topic_ViewCount          = arrTopicData(tT_VIEW_COUNT, iTopic)
		Topic_MemberID           = arrTopicData(mMEMBER_ID, iTopic)
		Topic_MemberName         = arrTopicData(mM_NAME, iTopic)
		Topic_LastPostAuthorName = arrTopicData(tLAST_POST_AUTHOR_NAME, iTopic)

		Dim AdminAllowed, ModerateAllowed
		if mLev = 4 then AdminAllowed = 1 else AdminAllowed = 0
		if mLev = 4 then
			ModerateAllowed = "Y"
		elseif mLev = 3 and ModOfForums <> "" then
			if (strAuthType = "nt") then
				if (chkForumModerator(Forum_ID, Session(strCookieURL & "username")) = "1") then ModerateAllowed = "Y" else ModerateAllowed = "N"
			else
				if (instr("," & ModOfForums & "," ,"," & Forum_ID & ",") <> 0) then ModerateAllowed = "Y" else ModerateAllowed = "N"
			end if
		else
			ModerateAllowed = "N"
		end if
		if ModerateAllowed = "Y" and Topic_UReplies > 0 then Topic_Replies = Topic_Replies + Topic_UReplies
		if chkDisplayForum(Forum_PrivateForums,Forum_FPasswordNew,Forum_ID,MemberID) then
			if (currForum <> Forum_ID) and (currTopic <> Topic_ID) then
				Response.Write "<tr>" & strLE & _
					"<td class=""ccc vam"" height=""20"" colspan=""6""><span class=""dff dfs cfc""><a href=""default.asp?CAT_ID=" & Forum_Cat_ID & """><span class=""dff dfs cfc""><b>" & ChkString(Cat_Name,"display") & "</b></span></a>&nbsp;/&nbsp;<a href=""forum.asp?FORUM_ID=" & Forum_ID & """><span class=""dff dfs cfc""><b>" & ChkString(Forum_Subject,"display") & "</b></span></a></td>" & strLE
				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
					Response.Write "<td class=""ccc vam c""><span class=""dff dfs cfc"">" & strLE
					Call ForumAdminOptions()
					Response.Write "</span></td>" & strLE
  				elseif (mLev = 3) then
					Response.Write "<td class=""ccc vam c""><span class=""dff dfs cfc""><b>&nbsp;</b></span></td>" & strLE
				end if
				Response.Write "</tr>" & strLE
				currForum = Forum_ID
			end if
			if currTopic <> Topic_ID then
				Response.Write "<tr>" & strLE
				if Cat_Status <> 0 and Forum_Status <> 0 and Topic_Status <> 0 then
					' DEM --> Added if statement to display topic status properly
					if Topic_Status = 2 then
						UnApprovedFound = "Y"
						Response.Write "<td class=""fcc c""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderUnmoderated,"Topic UnModerated","class=""vam""") & "</a></td>" & strLE
					elseif Topic_Status = 3 then
						HeldFound = "Y"
						Response.Write "<td class=""fcc c""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderHold,"Topic Held","class=""vam""") & "</a></td>" & strLE
					else
						Response.Write "<td class=""fcc c""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkIsNew(Topic_LastPost) & "</a></td>" & strLE
					end if
				else
 					if Cat_Status = 0 then
 						strAltText = "Category Locked"
 					elseif Forum_Status = 0 then
 						strAltText = "Forum Locked"
 					else
 						strAltText = "Topic Locked"
 					end if
 					Response.Write "<td class=""fcc c""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & getCurrentIcon(strIconFolderLocked,strAltText,"class=""vam""") & "</a></td>" & strLE
 				end if
 				Response.Write "<td class=""fcc""><span class=""dff dfs"">" & strLE & _
 					"<span class=""smt""><a href=""topic.asp?TOPIC_ID=" & Topic_ID & """>" & ChkString(left(Topic_Subject, 50),"display") & "</a></span>&nbsp;</span>" & strLE
 				if strShowPaging = "1" then TopicPaging()
 				Response.Write "</td>" & strLE & _
					"<td class=""fcc vat c""><span class=""dff dfs ffc""><span class=""smt"">" & profileLink(chkString(Topic_MemberName,"display"),Topic_Author) & "</span></span></td>" & strLE & _
					"<td class=""fcc vat c""><span class=""dff dfs ffc"">" & Topic_Replies & "</span></td>" & strLE & _
					"<td class=""fcc vat c""><span class=""dff dfs ffc"">" & Topic_ViewCount & "</span></td>" & strLE
 				if IsNull(Topic_LastPostAuthor) then
 					strLastAuthor = ""
 				else
 					strLastAuthor = "<br>by: <span class=""smt"">" & profileLink(Topic_LastPostAuthorName,Topic_LastPostAuthor) & "</span>"
 				end if
 				Response.Write "<td class=""fcc vam nw c""><span class=""dff ffc ffs""><b>" & ChkDate(Topic_LastPost, "</b>&nbsp" ,true) & strLastAuthor & "</span></td>" & strLE
 				if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
 					Response.Write "<td class=""fcc vam nw c""><span class=""dff dfs ffc"">" & strLE
 					call TopicAdminOptions
 					Response.Write "</span></td>" & strLE
 				elseif (mlev = 3) then
 					Response.Write "<td class=""fcc vam nw c""><span class=""dff dfs ffc"">&nbsp;</span></td>" & strLE
 				end if
 				Response.Write "</tr>" & strLE
 				currTopic = Topic_ID
 				rec = rec + 1
 			end if
 			mdisplayed = mdisplayed + 1
 		end if
 	next
 	if mdisplayed = 0 then
 		Response.Write "<tr>" & strLE & _
 			"<td class=""fcc"" colspan=""6""><span class=""dff dfs""><b>No Matches Found</b></span></td>" & strLE
 		if (mlev = 4 or mlev = 3) or (lcase(strNoCookies) = "1") then Response.Write "<td class=""fcc c""><b><span class=""dff dfs hfc"">&nbsp;</span></b></td>" & strLE
 		Response.Write "</tr>" & strLE
 	end if
 end if
 Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
 	"</table>" & strLE

if maxpages > 1 then
	Response.Write "<table class=""tc"" width=""100%"">" & strLE & _
		"<tr>" & strLE
	Call DropDownPaging(2)
	Response.Write "</tr>" & strLE & _
		"</table>" & strLE
end if

Response.Write "<table class=""tc"" width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""vat l"">" & strLE & _
	"<table>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw"">" & strLE & _
	"<p><span class=""dff ffs"">" & strLE & _
	getCurrentIcon(strIconFolderNew,"New Posts","class=""vam""") & " New posts since last logon.<br>" & strLE & _
	getCurrentIcon(strIconFolder,"Old Posts","class=""vam""") & " Old Posts."
if lcase(strHotTopic) = "1" then Response.Write (" (" & getCurrentIcon(strIconFolderHot,"Hot Topic","class=""vam""") & "&nbsp;" & intHotTopicNum & " replies or more.)<br>" & strLE)
Response.Write getCurrentIcon(strIconFolderLocked,"Locked Topic","class=""vam""") & " Locked topic.<br>" & strLE
' DEM --> Start of Code added for moderation
if HeldFound = "Y" then Response.Write getCurrentIcon(strIconFolderHold,"Held Posts","class=""vam""") & " Held Posts.<br>" & strLE
if UnApprovedFound = "Y" then Response.Write getCurrentIcon(strIconFolderUnmoderated,"UnModerated Posts","class=""vam""") & " UnModerated Posts.<br>" & strLE
' DEM --> End of Code added for moderation
Response.Write "</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
Call WriteFooter
Response.End

sub ForumAdminOptions()
	if (ModerateAllowed = "Y") or (lcase(strNoCookies) = "1") then
		' DEM --> Added code to allow for moderation
		if UnModeratedPosts > 0 then
			if ModerateAllowed = "Y" and (CheckForUnModeratedPosts("FORUM", Forum_Cat_ID, Forum_ID, 0) > 0) then
				ModString = "CAT_ID=" & Forum_Cat_ID & "&FORUM_ID=" & Forum_ID
				Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & ModString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts in this Forum","") & "</a>" & strLE
				UnModeratedFPosts = 1
			end if
		end if
		' DEM --> End of code added to allow for moderation
	end if
end sub

sub TopicPaging()
	mxpages = (Topic_Replies / strPageSize)
	if mxPages <> cLng(mxPages) then mxpages = int(mxpages) + 1
	if mxpages > 1 then
		Response.Write("<table cellspacing=""0"" cellpadding=""0"">" & strLE)
		Response.Write("<tr>" & strLE)
		Response.Write("<td class=""vam""><span class=""dff ffs"">" & getCurrentIcon(strIconFolderPosticon,"","") & "</span></td>" & strLE)
		for counter = 1 to mxpages
			ref = "<td class=""fcc vam r""><span class=""dff ffs"">"
			if ((mxpages > 9) and (mxpages > strPageNumberSize)) or ((counter > 9) and (mxpages < strPageNumberSize)) then ref = ref & "&nbsp;"
			ref = ref & widenum(counter) & "<span class=""smt""><a href=""topic.asp?"
			ref = ref & ArchiveLink
			ref = ref & "TOPIC_ID=" & Topic_ID
			ref = ref & "&whichpage=" & counter
			ref = ref & """>" & counter & "</a></span></span></td>"
			Response.Write ref & strLE
			if counter mod strPageNumberSize = 0 then
				Response.Write("</tr>" & strLE)
				Response.Write("<tr>" & strLE)
				Response.Write("<td>&nbsp;</td>" & strLE)
			end if
		next
		Response.Write("</tr>" & strLE)
		Response.Write("</table>" & strLE)
	end if
end sub

sub TopicAdminOptions()
	' DEM --> Start of Code for Full Moderation
	if UnModeratedFPosts > 0 then
        if CheckForUnModeratedPosts("TOPIC", Forum_Cat_ID, Forum_ID, Topic_ID) > 0 then
			TopicString = "TOPIC_ID=" & Topic_ID & "&CAT_ID=" & Forum_Cat_ID & "&FORUM_ID=" & Forum_ID
            Response.Write "<a href=""JavaScript:openWindow('pop_moderate.asp?" & TopicString & "')"">" & getCurrentIcon(strIconFolderModerate,"Approve/Hold/Reject All UnModerated Posts for this Topic","class=""vam""") & "</a>" & strLE
        end if
	end if
	' DEM --> End of Code for Full Moderation
end sub

sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then pge = 1 else pge = mypage
		Response.Write "<form name=""PageNum" & fnum & """ action=""moderate.asp"" method=""post"">" & strLE
		Response.Write "<td><span class=""dff dfs"">" & strLE
		if strArchiveState = "1" and ArchiveView = "true" then Response.Write("<input name=""ARCHIVE"" type=""hidden"" value=""" & ArchiveView & """>" & strLE)
		if fnum = 1 then
			Response.Write("<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
		else
			Response.Write("<b>There are " & maxpages & " Pages of Unmoderated Posts: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		if fnum = 1 then
			Response.Write("</select><b> of " & maxPages & "</b>" & strLE)
		else
			Response.Write("</select>" & strLE)
		end if
		Response.Write("</span></td>" & strLE)
		Response.Write("</form>" & strLE)
	end if
end sub
%>
