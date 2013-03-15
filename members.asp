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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if strDBNTUserName = "" then
	Err_Msg = "<li>You must be logged in to view the Members List</li>"
	Response.Write "<table width=""100%"">" & strLE & _
		"<tr>" & strLE & _
		"<td><span class=""dff dfs"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member Information</span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem!</span></p>" & strLE & _
		"<p class=""c""><span class=""dff dfs hlfc"">You must be logged in to view this page</span></p>" & strLE & _
		"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Back to Forum</a></span></p>" & strLE & _
		"<br>" & strLE
	Call WriteFooter
	Response.End
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

if trim(chkString(Request("method"),"SQLString")) <> "" then
	SortMethod     = trim(chkString(Request("method"),"SQLString"))
	strSortMethod  = "&method=" & SortMethod
	strSortMethod2 = "?method=" & SortMethod
end if

if trim(chkString(Request("mode"),"SQLString")) <> "" then
	strMode = trim(chkString(Request("mode"),"SQLString"))
	if strMode <> "search" then strMode = ""
end if

SearchName = trim(Request("M_NAME"))
if SearchName = "" then SearchName = trim(Request.Form("M_NAME"))
SearchNameDisplay = Server.HTMLEncode(SearchName)
If SearchName <> "" Then
	If Not IsValidString(SearchName) Then
		Err_Msg = "Invalid Name!"
		Response.Write "<table width=""100%"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span class=""dff dfs"">" & strLE & _
			getCurrentIcon(strIconFolderOpen,"","") & " <a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
			getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & " Member Information</span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs hlfc"">" & Err_Msg & "</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Back to Forum</a></span></p>" & strLE & _
			"<br>" & strLE
	Call WriteFooter
	Response.End
	End If
End if
SearchName = chkString(SearchName, "sqlstring")
if Request("UserName") <> "" then
	if IsNumeric(Request("UserName")) = True then srchUName = cLng(Request("UserName")) else srchUName = "1"
end if
if Request("FirstName") <> "" then
	if IsNumeric(Request("FirstName")) = True then srchFName = cLng(Request("FirstName")) else srchFName = "0"
end if
if Request("LastName") <> "" then
	if IsNumeric(Request("LastName")) = True then srchLName = cLng(Request("LastName")) else srchLName = "0"
end if
if Request("INITIAL") <> "" then
	if IsNumeric(Request("INITIAL")) = True then srchInitial = cLng(Request("INITIAL")) else srchInitial = "0"
end if

mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)

'New Search Code
If strMode = "search"  and (srchUName = "1" or srchFName = "1" or srchLName = "1" or srchInitial = "1" ) then
	strSql  = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql  = strSql & "M_AIM, M_ICQ, M_MSN, M_YAHOO, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE "
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
'	if Request.querystring("link") <> "sort" then
		whereSql = " WHERE ("
		tmpSql   = ""
		if srchUName = "1" then
			tmpSql = tmpSql & "M_NAME LIKE '%" & SearchName & "%' OR "
			tmpSql = tmpSql & "M_USERNAME LIKE '%" & SearchName & "%'"
		end if
		if srchFName = "1" then
			if srchUName = "1" then tmpSql = tmpSql & " OR "
			tmpSql = tmpSql & "M_FIRSTNAME LIKE '%" & SearchName & "%'"
		end if
		if srchLName = "1" then
			if srchFName = "1" or srchUName = "1" then tmpSql = tmpSql & " OR "
			tmpSql = tmpSql & "M_LASTNAME LIKE '%" & SearchName & "%' "
		end if
		if srchInitial = "1" then tmpSQL = "M_NAME LIKE '" & SearchName & "%'"
		whereSql = whereSql & tmpSql &")"
		Session(strCookieURL & "where_Sql") = whereSql
'	end if
	if Session(strCookieURL & "where_Sql") <> "" then
		whereSql = Session(strCookieURL & "where_Sql")
	else
		whereSql = ""
	end if
	strSQL3 = whereSql
else
	'## Forum_SQL - Get all members
	strSql  = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_COUNTRY, M_HOMEPAGE, "
	strSql  = strSql & "M_AIM, M_ICQ, M_MSN, M_YAHOO, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE "
	strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
	if mlev = 4 then
		strSql3 = " WHERE M_NAME <> 'n/a' "
	else
		strSql3 = " WHERE M_STATUS = " & 1
	end if
end if
select case SortMethod
	case "nameasc" : strSql4 = " ORDER BY M_NAME ASC"
	case "namedesc" : strSql4 = " ORDER BY M_NAME DESC"
	case "levelasc" : strSql4 = " ORDER BY M_TITLE ASC, M_NAME ASC"
	case "leveldesc" : strSql4 = " ORDER BY M_TITLE DESC, M_NAME ASC"
	case "lastpostdateasc" : strSql4 = " ORDER BY M_LASTPOSTDATE ASC, M_NAME ASC"
	case "lastpostdatedesc" : strSql4 = " ORDER BY M_LASTPOSTDATE DESC, M_NAME ASC"
	case "lastheredateasc"
		if mlev = 4 or mlev = 3 then
			strSql4 = " ORDER BY M_LASTHEREDATE ASC, M_NAME ASC"
		else
			strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
		end if
	case "lastheredatedesc"
		if mlev = 4 or mlev = 3 then
			strSql4 = " ORDER BY M_LASTHEREDATE DESC, M_NAME ASC"
		else
			strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
		end if
	case "dateasc" : strSql4 = " ORDER BY M_DATE ASC, M_NAME ASC"
	case "datedesc" : strSql4 = " ORDER BY M_DATE DESC, M_NAME ASC"
	case "countryasc" : strSql4 = " ORDER BY M_COUNTRY ASC, M_NAME ASC"
	case "countrydesc" : strSql4 = " ORDER BY M_COUNTRY DESC, M_NAME ASC"
	case "postsasc" : strSql4 = " ORDER BY M_POSTS ASC, M_NAME ASC"
	case else : strSql4 = " ORDER BY M_POSTS DESC, M_NAME ASC"
end select

if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		OffSet  = cLng((mypage - 1) * strPageSize)
		strSql5 = " LIMIT " & OffSet & ", " & strPageSize & " "
	end if
	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "
	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing
	if iPageTotal > 0 then
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
		maxpages = 0
	end if
	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		arrMemberData = rs.GetRows(intGetRows)
		iMemberCount  = UBound(arrMemberData, 2)
		rs.close
		set rs = nothing
	else
		iMemberCount = ""
	end if
else 'end MySql specific code
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		If not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize     = strPageSize
			rs.absolutepage = mypage '**
			maxpages        = cLng(rs.pagecount)
			arrMemberData   = rs.GetRows(strPageSize)
			iMemberCount    = UBound(arrMemberData, 2)
		else
			iMemberCount = ""
		end if
	rs.Close
	set rs = nothing
end if
Response.Write "<table width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member Information</span></td>" & strLE & _
	"<td class=""vab r"">" & strLE
if maxpages > 1 then
	Response.Write "<table class=""tc"">>" & strLE & _
		"<tr>" & strLE
	Call Paging2(1)
	Response.Write "</tr>" & strLE & _
		"</table>" & strLE
else
	Response.Write "&nbsp;" & strLE
end if
Response.Write "</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE

Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""pubc"">" & strLE & _
	"<table width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<form action=""members.asp" & strSortMethod2 & """ method=""post"" name=""SearchMembers"">" & strLE & _
	"<td class=""putc""><span class=""dff dfs""><b>Search:</b>&nbsp;" & strLE & _
	"<input type=""checkbox"" name=""UserName"" value=""1"""
if ((srchUName <> "")  or (srchUName = "" and srchFName = "" and srchLName = "") ) then Response.Write(" checked")
Response.Write ">User Names" & strLE
if strFullName = "1" then
	Response.Write "&nbsp;&nbsp;<input type=""checkbox"" name=""FirstName"" value=""1""" & chkCheckbox(srchFName,1,true) & ">First Name" & strLE & _
		"&nbsp;&nbsp;<input type=""checkbox"" name=""LastName"" value=""1""" & chkCheckbox(srchLName,1,true) & ">Last Name" & strLE
end if
Response.Write "</span></td>" & strLE & _
	"<td class=""putc""><span class=""dff dfs""><b>For:</b>&nbsp;" & strLE & _
	"<input type=""text"" name=""M_NAME"" value=""" & SearchNameDisplay & """></span></td>" & strLE & _
	"<input type=""hidden"" name=""mode"" value=""search"">" & strLE & _
	"<input type=""hidden"" name=""initial"" value=""0"">" & strLE & _
	"<td class=""putc c"">" & strLE
if strGfxButtons = "1" then
	'Response.Write "<input type=""submit"" value=""search"" style=""color:" & strPopUpBorderColor & ";border: 1px solid " & strPopUpBorderColor & "; background-color: " & strPopUpTableColor & "; cursor: hand;"" id=""submit1"" name=""submit1"">" & strLE
	Response.Write "<input src=""" & strImageUrl & "button_go.gif"" alt=""Quick Search"" type=""image"" value=""search"" id=""submit1"" name=""submit1"">" & strLE
else
	Response.Write "<input type=""submit"" value=""search"" id=""submit1"" name=""submit1"">" & strLE
end if
Response.Write "</td>" & strLE & _
	"</form>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""putc"">" & strLE & _
	"<td class=""vat c"" colspan=""3""><span class=""dff dfs"">" & strLE & _
	"<a href=""members.asp"">All</a>&nbsp;" & strLE
for intChar = 65 to 90
	if intChar <> 90 then
		Response.Write "<a href=""members.asp?mode=search&M_NAME=" & chr(intChar) & "&initial=1" & strSortMethod & """>" & chr(intChar) & "</a>&nbsp;" & strLE
	else
		Response.Write "<a href=""members.asp?mode=search&M_NAME=" & chr(intChar) & "&initial=1" & strSortMethod & """>" & chr(intChar) & "</a><br></span></td>" & strLE
	end if
next
Response.Write "</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE & _
	"<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""3"">" & strLE & _
	"<tr>" & strLE
strNames = "UserName=" & srchUName  &_
	"&FirstName=" & srchFName &_
	"&LastName=" & srchLName &_
	"&INITIAL=" &srchInitial & "&"

Response.Write "<td class=""hcc c""><b><span class=""dff dfs hfc"">&nbsp;&nbsp;</span></b></td>" & strLE & _
	"<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "nameasc" then Response.Write("namedesc") else Response.Write("nameasc")
Response.Write """><b><span class=""dff dfs hfc"">Member Name</span></b></a></td>" & strLE & _
	"<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "levelasc" then Response.Write("leveldesc") else Response.Write("levelasc")
Response.Write """><b><span class=""dff dfs hfc"">Title</span></b></a></td>" & strLE & _
	"<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "postsdesc" then Response.Write("postsasc") else Response.Write("postsdesc")
Response.Write """><b><span class=""dff dfs hfc"">Posts</span></b></a></td>" & strLE & _
	"<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "lastpostdatedesc" then Response.Write("lastpostdateasc") else Response.Write("lastpostdatedesc")
Response.Write """><b><span class=""dff dfs hfc"">Last Post</span></b></a></td>" & strLE & _
	"<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
if Request.QueryString("method") = "datedesc" then Response.Write("dateasc") else Response.Write("datedesc")
Response.Write """><b><span class=""dff dfs hfc"">Member Since</span></b></a></td>" & strLE
if strCountry = "1" then
	Response.Write "<td class=""hcc c""><a href=""members.asp?" & strNames & "link=sort&mode=search&M_NAME=" & SearchName & "&method="
	if Request.QueryString("method") = "countryasc" then Response.Write("countrydesc") else Response.Write("countryasc")
	Response.Write """><b><span class=""dff dfs hfc"">Country</span></b></a></td>" & strLE
end if
if mlev = 4 or mlev = 3 then
	Response.Write "<td class=""hcc c""><a href=""members.asp?method="
	if Request.QueryString("method") = "lastheredatedesc" then Response.Write("lastheredateasc") else Response.Write("lastheredatedesc")
	Response.Write """><b><span class=""dff dfs hfc"">Last Visit</span></b></a></td>" & strLE
end if
if mlev = 4 or (lcase(strNoCookies) = "1") then Response.Write "<td class=""hcc c""><b><span class=""dff dfs hfc"">&nbsp;</span></b></td>" & strLE
Response.Write "</tr>" & strLE
if iMemberCount = "" then '## No Members Found in DB
	Response.Write "<tr>" & strLE & _
		"<td colspan=""" & sGetColspan(9, 8) & """ class=""fcc""><span class=""dff dfs ffc""><b>No Members Found</b></span></td>" & strLE & _
		"</tr>" & strLE
else
	mMEMBER_ID      = 0
	mM_STATUS       = 1
	mM_NAME         = 2
	mM_LEVEL        = 3
	mM_EMAIL        = 4
	mM_COUNTRY      = 5
	mM_HOMEPAGE     = 6
	mM_AIM          = 7
	mM_ICQ          = 8
	mM_MSN          = 9
	mM_YAHOO        = 10
	mM_TITLE        = 11
	mM_POSTS        = 12
	mM_LASTPOSTDATE = 13
	mM_LASTHEREDATE = 14
	mM_DATE         = 15

	rec  = 1
	intI = 0
	for iMember = 0 to iMemberCount
		if (rec = strPageSize + 1) then exit for

		Members_MemberID           = arrMemberData(mMEMBER_ID, iMember)
		Members_MemberStatus       = arrMemberData(mM_STATUS, iMember)
		Members_MemberName         = arrMemberData(mM_NAME, iMember)
		Members_MemberLevel        = arrMemberData(mM_LEVEL, iMember)
		Members_MemberEMail        = arrMemberData(mM_EMAIL, iMember)
		Members_MemberCountry      = arrMemberData(mM_COUNTRY, iMember)
		Members_MemberHomepage     = arrMemberData(mM_HOMEPAGE, iMember)
		Members_MemberAIM          = arrMemberData(mM_AIM, iMember)
		Members_MemberICQ          = arrMemberData(mM_ICQ, iMember)
		Members_MemberMSN          = arrMemberData(mM_MSN, iMember)
		Members_MemberYAHOO        = arrMemberData(mM_YAHOO, iMember)
		Members_MemberTitle        = arrMemberData(mM_TITLE, iMember)
		Members_MemberPosts        = arrMemberData(mM_POSTS, iMember)
		Members_MemberLastPostDate = arrMemberData(mM_LASTPOSTDATE, iMember)
		Members_MemberLastHereDate = arrMemberData(mM_LASTHEREDATE, iMember)
		Members_MemberDate         = arrMemberData(mM_DATE, iMember)

		if intI = 1 then
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if

		Response.Write "<tr>" & strLE & _
			"<td class=""c"" bgcolor=""" & CColor & """>" & strLE
		if strUseExtendedProfile then
			Response.Write "<a href=""pop_profile.asp?mode=display&id=" & Members_MemberID & """>"
		else
			Response.Write "<a href=""JavaScript:openWindow3('pop_profile.asp?mode=display&id=" & Members_MemberID & "')"">"
		end if
		if Members_MemberStatus = 0 then
			Response.Write getCurrentIcon(strIconProfileLocked,"View " & ChkString(Members_MemberName,"display") & "'s Profile","class=""vam""")
		else
			Response.Write getCurrentIcon(strIconProfile,"View " & ChkString(Members_MemberName,"display") & "'s Profile","class=""vam""")
		end if
		Response.Write "</a>" & strLE
		if strAIM = "1" and Trim(Members_MemberAIM) <> "" then Response.Write "<a href=""JavaScript:openWindow('pop_messengers.asp?mode=AIM&ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconAIM,"Send " & ChkString(Members_MemberName,"display") & " an AOL message","class=""vam""") & "</a>" & strLE
		if strICQ = "1" and Trim(Members_MemberICQ) <> "" then Response.Write "<a href=""JavaScript:openWindow6('pop_messengers.asp?mode=ICQ&ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconICQ,"Send " & ChkString(Members_MemberName,"display") & " an ICQ Message","class=""vam""") & "</a>" & strLE
		if strMSN = "1" and Trim(Members_MemberMSN) <> "" then Response.Write "<a href=""JavaScript:openWindow('pop_messengers.asp?mode=MSN&ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconMSNM,"Click to see " & ChkString(Members_MemberName,"display") & "'s MSN Messenger address","class=""vam""") & "</a>" & strLE
		if strYAHOO = "1" and Trim(Members_MemberYAHOO) <> "" then Response.Write "<a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(Members_MemberYAHOO, "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon(strIconYahoo,"Send " & ChkString(Members_MemberName,"display") & " a Yahoo! Message","class=""vam""") & "</a>" & strLE
		Response.Write "</td>" & strLE & _
			"<td bgcolor=""" & CColor & """><span class=""dff dfs"">" & strLE
		if strUseExtendedProfile then
			Response.Write "<span class=""smt""><a href=""pop_profile.asp?mode=display&id=" & Members_MemberID & """ title=""View " & ChkString(Members_MemberName,"display") & "'s Profile"">"
		else
			Response.Write "<span class=""smt""><a href=""JavaScript:openWindow3('pop_profile.asp?mode=display&id=" & Members_MemberID & "')"" title=""View " & ChkString(Members_MemberName,"display") & "'s Profile"">"
		end if
		Response.Write ChkString(Members_MemberName,"display") & "</a></span></span></td>" & strLE & _
			"<td class=""c"" bgcolor=""" & CColor & """><span class=""dff dfs ffc"">" & ChkString(getMember_Level(Members_MemberTitle, Members_MemberLevel, Members_MemberPosts),"display") & "</span></td>" & strLE & _
			"<td class=""c"" bgcolor=""" & CColor & """><span class=""dff dfs ffc"">"
		if IsNull(Members_MemberPosts) then
			Response.Write("-")
		else
			Response.Write(Members_MemberPosts)
			if strShowRank = 2 or strShowRank = 3 then Response.Write("<br>" & getStar_Level(Members_MemberLevel, Members_MemberPosts) & "")
		end if
		Response.Write "</span></td>" & strLE
		if IsNull(Members_MemberLastPostDate) or Trim(Members_MemberLastPostDate) = "" then
			Response.Write "<td bgcolor=""" & CColor & """ class=""nw c""><span class=""dff dfs ffc"">-</span></td>" & strLE
		else
			Response.Write "<td bgcolor=""" & CColor & """ class=""nw c""><span class=""dff dfs ffc"">" & ChkDate(Members_MemberLastPostDate,"",false) & "</span></td>" & strLE
		end if
		Response.Write "<td bgcolor=""" & CColor & """ class=""nw c""><span class=""dff dfs ffc"">" & ChkDate(Members_MemberDate,"",false) & "</span></td>" & strLE
		if strCountry = "1" then
			Response.Write "<td class=""c"" bgcolor=""" & CColor & """><span class=""dff dfs ffc"">"
			if trim(Members_MemberCountry) <> "" then Response.Write(Members_MemberCountry & "&nbsp;") else Response.Write("-")
			Response.Write "</span></td>" & strLE
		end if
		if mlev = 4 or mlev = 3 then Response.Write "<td bgcolor=""" & CColor & """ class=""nw c""><span class=""dff dfs ffc"">" & ChkDate(Members_MemberLastHereDate,"",false) & "</span></td>" & strLE
		if mlev = 4 or (lcase(strNoCookies) = "1") then
			Response.Write "<td class=""c"" bgcolor=""" & CColor & """><b><span class=""dff dfs"">" & strLE
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				if Members_MemberStatus <> 0 then
					Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconLock,"Lock Member","class=""vam""") & "</a>" & strLE
					Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Zap&MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconZap,"Zap Member Profile","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (Members_MemberID = intAdminMemberID and MemberID <> intAdminMemberID) OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID AND MemberID <> Members_MemberID) then
				Response.Write "                -" & strLE
			else
				if strUseExtendedProfile then
					Response.Write "<a href=""pop_profile.asp?mode=Modify&ID=" & Members_MemberID & """>" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow3('pop_profile.asp?mode=Modify&ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?mode=Member&MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Member","class=""vam""") & "</a>" & strLE
			end if
			Response.Write "</span></b></td>" & strLE
		end if
		Response.Write "</tr>" & strLE
		rec = rec + 1
		intI = intI + 1
		if intI = 2 then intI = 0
	next
end if
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td colspan=""2"">" & strLE
if maxpages > 1 then
	Response.Write "<table>" & strLE & _
		"<tr>" & strLE
	Call Paging2(2)
	Response.Write "</tr>" & strLE & _
		"</table>" & strLE
end if
Response.Write "</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE
Call WriteFooter
Response.End

sub Paging2(fnum)
	if maxpages > 1 then
		if mypage = "" then sPageNumber = 1 else sPageNumber = mypage
		if SortMethod = "" then sMethod = "postsdesc" else sMethod = SortMethod
		Response.Write("<form name=""PageNum" & fnum & """ action=""members.asp"">" & strLE)
		if fnum = 1 then
			Response.Write("<td class=""vab r""><span class=""dff dfs"">" & strLE)
		else
			Response.Write("<td><span class=""dff dfs"">" & strLE)
		end if
		if srchInitial <> "" then Response.Write("<input type=""hidden"" name=""initial"" value=""" & srchInitial & """>" & strLE)
		if sMethod <> "" then Response.Write("<input type=""hidden"" name=""method"" value=""" & sMethod & """>" & strLE)
		if strMode <> "" then Response.Write("<input type=""hidden"" name=""mode"" value=""" & strMode & """>" & strLE)
		if searchName <> "" then Response.Write("<input type=""hidden"" name=""M_NAME"" value=""" & searchName & """>" & strLE)
		if srchUName <> "" then Response.write("<input type=""hidden"" name=""UserName"" value=""" & srchUName & """>" & strLE)
		if srchFName <> "" then Response.write("<input type=""hidden"" name=""FirstName"" value=""" & srchFName & """>" & strLE)
		if srchLName <> "" then Response.write("<input type=""hidden"" name=""LastName"" value=""" & srchLName & """>" & strLE)
		if fnum = 1 then
			Response.Write("<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
        	else
			Response.Write("<b>Members are " & maxpages & " Pages Long: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(sPageNumber) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		Response.Write "</select>"
		if fnum = 1 then Response.Write "<b> of " & maxPages & "</b>" & strLE
		Response.Write("</span></td>" & strLE)
		Response.Write("</form>" & strLE)
	end if
end sub

Function sGetColspan(lIN, lOUT)
	if (mlev = "4" or mlev = "3") then lOut = lOut + 2
	if lOut > lIn then sGetColspan = lIN Else sGetColspan = lOUT
end Function
%>
