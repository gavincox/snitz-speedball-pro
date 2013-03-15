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

'#################################################################################
'## Subscription_List.asp - This page will search through all subscriptions.
'##                         If the user is an administrator, then it will loop
'##                         through all the subscriptions, otherwise it will only
'##                         look for those subscriptions which apply directly to
'##                         them.
'#################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_subscription.asp"-->
<!--#INCLUDE FILE="cb/subscription_list_cb.asp" -->
<%
' -- Make sure user is logged on.
if strDBNTUserName = "" then
	Response.redirect ("default.asp")
else
	' -- ensure that only admin's can look at ALL subscriptions.
	If mlev <> 4 then Mode = "" else Mode = Request("MODE")
	' -- display the appropriate message
	if Mode = "" then
		strPageTitle = "Subscriptions for <b>" & strDBNTUserName & "</b>"
	else
		strPageTitle = "Subscriptions for <b>All Members</b>"
	end if
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs w50"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & " <a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE  & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & " " & strPageTitle & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE
dim HeadBg, HeadFont, CellBg, CellFont, ForumBg, ForumFont, LoopCount, TotalSubCount
' -- Added to make CSS easier hopefully
HeadBg 		= "class=""hcc"""
HeadFont	= "<span class=""dff dfs hfc"">"
ForumBg		= "bgCategoryCellColor"
ForumFont	= "<span class=""dff dfs cfc"">"
CellBg		= "bgForumCellColor "
CellFont	= "<span class=""dff dfs ffc"">"
AltCellBg	= "bgAltForumCellColor "
' If no subscriptions allowed - exit
if strSubscription = 0 then Go_Result : Response.End
' Look for all applicable subscriptions.....
StrSQL = "SELECT S.SUBSCRIPTION_ID, S.MEMBER_ID, M.M_NAME," & _
	"S.CAT_ID, C.CAT_NAME, C.CAT_STATUS, C.CAT_SUBSCRIPTION, " & _
	"S.FORUM_ID, F.F_SUBJECT, F.F_STATUS, F.F_SUBSCRIPTION, " & _
	"S.TOPIC_ID, T.T_SUBJECT, T.T_STATUS " & _
	"FROM (((" & strTablePrefix & "SUBSCRIPTIONS S INNER JOIN " & strMemberTablePrefix & "MEMBERS M ON S.MEMBER_ID = M.MEMBER_ID) " & _
	"LEFT JOIN " & strTablePrefix & "TOPICS T ON S.TOPIC_ID = T.TOPIC_ID) " & _
	"LEFT JOIN " & strTablePrefix & "FORUM F ON S.FORUM_ID = F.FORUM_ID) " & _
	"LEFT JOIN " & strTablePrefix & "CATEGORY C ON S.CAT_ID = C.CAT_ID "
if Mode = "" then strSQL = strSQL & "WHERE S.MEMBER_ID = " & MemberID & " "
strSQL = strSQL & "ORDER BY C.CAT_ORDER, C.CAT_NAME, F.F_ORDER, F.F_SUBJECT, S.TOPIC_ID ASC"
set rs = Server.CreateObject("ADODB.Recordset")
rs.Open StrSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
if rs.EOF or rs.BOF then
	' -- If none are found, exit...
	Go_Result
else
	Response.Write "<table id=""content"">" & strLE
	HldCatID = -99 : HldForumID = -99 : HldTopicID = -99 ' Used for displaying titles...
	arrSubs  = rs.GetRows(adGetRowsRest)
	SubCount = UBound(arrSubs, 2)
	rs.Close
	set rs = nothing
	iSubCount = 0
		for isub = 0 to SubCount
		iSubCount         = iSubCount + 1
		' -- Move values from the array to local variables...
		SubscriptionID    = arrSubs(0,isub)
		SubMemberID       = arrSubs(1,isub)
		SubMemberName     = arrSubs(2,isub)
		CatID             = cLng(arrSubs(3,isub))
		CatStatus         = arrSubs(5,isub)
		CatName           = arrSubs(4,isub)
		CatSubscription   = arrSubs(6,isub)
		ForumID           = cLng(arrSubs(7,isub))
		ForumStatus       = arrSubs(9,isub)
		ForumSubject      = arrSubs(8,isub)
		ForumSubscription = arrSubs(10,isub)
		TopicID           = cLng(arrSubs(11,isub))
		TopicStatus       = arrSubs(13,isub)
		TopicSubject      = arrSubs(12, isub)
		if CatID <> HldCatID then
			if CatID = 0 then
				DisplayText = "Board Level Subscriptions" & GetSubLevel(strSubscription)
				HldForumID  = 0 : HldTopicID = 0
			else
				DisplayText = "Category: " & CatName & GetSubLevel(CatSubscription)
				HldForumID  = -99 : HldTopicID = -99
			end if
			Response.Write "<tr>" & strLE & _
				"<th colspan=""2"">" & DisplayText & "</th>" & strLE & _
				"</tr>" & strLE
			HldCatID = CatID
		end if
		if ForumID <> HldForumID then
			if ForumID = 0 then
				DisplayText = "Category Level Subscriptions" : HldTopicID = 0
			else
				DisplayText = "Forum: " & ForumSubject	& GetFSubLevel(ForumSubscription)
				HldTopicID  = -99
			end if
			Response.Write "<tr>" & strLE & _
				"<th colspan=""2"">" & getCurrentIcon(strIconGoRight,"","") & DisplayText & "</th>" & strLE & _
				"</tr>" & strLE
			HldForumID = ForumID
		end if
		if TopicID <> HldTopicID then
			if TopicID = 0 then
				DisplayText = "Forum Level Subscriptions"
			else
				DisplayText = "<b>Topic: </b><span class=""smt""><a href=""topic.asp?TOPIC_ID=" & TopicID & """>" & TopicSubject & "</a>"
			end if
			Response.Write "<tr>" & strLE & _
				"<th colspan=""2"">" & getCurrentIcon(strIconGoRight,"","") & getCurrentIcon(strIconGoRight,"","") & DisplayText & "</th>" & strLE & _
				"</tr>" & strLE
			HldTopicID = TopicID
		end if
		LinkStartText = "<a href=""Javascript:unsub_confirm('pop_subscription.asp?subscribe=U&amp;MEMBER_ID=" & SubMemberID & "&amp;LEVEL="
		LinkEndText = "')"">" & getCurrentIcon(strIconUnsubscribe,"Unsubscribe","class=""vam""") & "</a>"
		Response.Write "<tr>" & strLE & _
			"<td class=""w95 vam"">"
		if CatID = 0 then
			Response.Write getCurrentIcon(strIconGoRight,"","class=""vam""")
			LinkText = "BOARD"
		elseif ForumID = 0 then
			Response.Write getCurrentIcon(strIconGoRight,"","class=""vam""") & getCurrentIcon(strIconGoRight,"","class=""vam""")
			LinkText = "CAT&amp;CAT_ID=" & CatID
		else
			Response.Write getCurrentIcon(strIconGoRight,"","class=""vam""") & getCurrentIcon(strIconGoRight,"","class=""vam""") & getCurrentIcon(strIconGoRight,"","class=""vam""")
			if TopicID = 0 then
				LinkText = "FORUM&amp;CAT_ID=" & CatID & "&amp;FORUM_ID=" & ForumID
			else
				LinkText = "TOPIC&amp;CAT_ID=" & CatID & "&amp;FORUM_ID=" & ForumID & "&amp;TOPIC_ID=" & TopicID
			end if
		end if
		Response.Write " " & SubMemberName & "</span></td>" & strLE & _
			"<td class=""c"">" & LinkStartText & LinkText & LinkEndText & "</td>" & strLE & _
			"</tr>" & strLE
	next
end if
Call Go_Result
%>
