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
<!--#INCLUDE FILE="inc_func_common.asp" -->
<%

if strShowTimer = "1" then
	'### start of timer code
	Dim StopWatch(19)

	sub StartTimer(x)
		StopWatch(x) = timer
	end sub

	function StopTimer(x)
		EndTime = Timer

		'Watch for the midnight wraparound...
		if EndTime < StopWatch(x) then
			EndTime = EndTime + (86400)
		end if

		StopTimer = EndTime - StopWatch(x)
	end function

	StartTimer 1

	'### end of timer code
end if

strArchiveTablePrefix = strTablePrefix & "A_"
strScriptName = request.servervariables("script_name")
strReferer = chkString(request.servervariables("HTTP_REFERER"),"refer")

if Application(strCookieURL & "down") then
	if not Instr(strScriptName,"admin_") > 0 then
		Response.redirect("down.asp")
	end if
end if

if strPageBGImageURL = "" then
	strTmpPageBGImageURL = ""
elseif Instr(strPageBGImageURL,"/") > 0 or Instr(strPageBGImageURL,"\") > 0 then
	strTmpPageBGImageURL = " background=""" & strPageBGImageURL & """"
else
	strTmpPageBGImageURL = " background=""" & strImageUrl & strPageBGImageURL & """"
end if

If strDBType = "" then
	Response.Write "<!doctype html>" & strLE & _
		"<html lang=""en"">" & strLE & _
		"<head>" & strLE & _
		"<title>" & chkString(strForumTitle,"pagetitle") & "</title>" & strLE
'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<meta property=""dcterms.rights"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write _
		"<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & strLE & _
		"<link href=""css/normalize-legacy.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"<link href=""css/snitz.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"</head>" & strLE & _
		"<body " & strTmpPageBGImageURL & ">" & strLE & _
		"<table class=""tc"" style=""padding:5px;width:50%;height:40%"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" bgColor=""#9FAFDF""><p><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
		"<b>There has been a problem...</b><br><br>Your <b>strDBType</b> is not set, please edit your <b>config.asp</b><br>to reflect your database type." & _
		"</font></p></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c""><font face=""Verdana, Arial, Helvetica"" size=""2"">" & _
		"<a href=""default.asp"" target=""_top"">Click here to retry.</a></font></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</body>" & strLE & _
		"</html>" & strLE
	Response.End
end if

set my_Conn = Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

if (strAuthType = "nt") then
	call NTauthenticate()
	if (ChkAccountReg() = "1") then
		call NTUser()
	end if
end if

if strGroupCategories = "1" then
	if Request.QueryString("Group") = "" then
		if Request.Cookies(strCookieURL & "GROUP") = "" Then
			Group = 2
		else
			Group = cLng(Request.Cookies(strCookieURL & "GROUP"))
		end if
	else
		Group = cLng(Request.QueryString("Group"))
	end if
	'set default
	Session(strCookieURL & "GROUP_ICON") = "icon_group_categories.gif"
	Session(strCookieURL & "GROUP_IMAGE") = strTitleImage
	'Forum_SQL - Group exists ?
	strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE "
	strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rs2 = my_Conn.Execute (strSql)
	if rs2.EOF or rs2.BOF then
		Group = 2
		strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_ICON, GROUP_IMAGE "
		strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
		strSql = strSql & " WHERE GROUP_ID = " & Group
		set rs2 = my_Conn.Execute (strSql)
	end if
	Session(strCookieURL & "GROUP_NAME") = rs2("GROUP_NAME")
	if instr(rs2("GROUP_ICON"), ".") then
		Session(strCookieURL & "GROUP_ICON") = rs2("GROUP_ICON")
	end if
	if instr(rs2("GROUP_IMAGE"), ".") then
		Session(strCookieURL & "GROUP_IMAGE") = rs2("GROUP_IMAGE")
	end if
	rs2.Close
	set rs2 = nothing
	Response.Cookies(strCookieURL & "GROUP") = Group
	Response.Cookies(strCookieURL & "GROUP").Expires =  dateAdd("d", intCookieDuration, strForumTimeAdjust)
	if Session(strCookieURL & "GROUP_IMAGE") <> "" then
		strTitleImage = Session(strCookieURL & "GROUP_IMAGE")
	end if
end if

strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
strDBNTFUserName = trim(chkString(Request.Form("Name"),"SQLString"))
if strDBNTFUserName = "" then strDBNTFUserName = trim(chkString(Request.Form("User"),"SQLString"))
if strAuthType = "nt" then
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
end if

if strRequireReg = "1" and strDBNTUserName = "" then
	if not Instr(strScriptName,"register.asp") > 0 and _
	not Instr(strScriptName,"password.asp") > 0 and _
	not Instr(strScriptName,"faq.asp") > 0 and _
	not Instr(strScriptName,"login.asp") > 0 then
		scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
		if Request.QueryString <> "" then
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))) & "?" & Request.QueryString)
		else
			Response.Redirect("login.asp?target=" & lcase(scriptname(ubound(scriptname))))
		end if
	end if
end if

select case Request.Form("Method_Type")
	case "login"
		strEncodedPassword = sha256("" & Request.Form("Password"))
		select case chkUser(strDBNTFUserName, strEncodedPassword,-1)
			case 1, 2, 3, 4
				Call DoCookies(Request.Form("SavePassword"))
				strLoginStatus = 1
			case else
				strLoginStatus = 0
			end select
	case "logout"
		Call ClearCookies()
end select

if trim(strDBNTUserName) <> "" and trim(Request.Cookies(strUniqueID & "User")("Pword")) <> "" then
	chkCookie = 1
	mLev = cLng(chkUser(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"),-1))
	chkCookie = 0
else
	MemberID = -1
	mLev = 0
end if

if mLev = 4 and strEmailVal = "1" and strRestrictReg = "1" and strEmail = "1" then
	'## Forum_SQL - Get membercount from DB
	strSql = "SELECT COUNT(MEMBER_ID) AS U_COUNT FROM " & strMemberTablePrefix & "MEMBERS_PENDING WHERE M_APPROVE = " & 0

	set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if not rs.EOF then
		User_Count = cLng(rs("U_COUNT"))
	else
		User_Count = 0
	end if

	rs.close
	set rs = nothing
end if

Response.Write "<!doctype html>" & strLE & _
	"<html lang=""en"">" & strLE & strLE & _
	"<head>" & strLE & _
	"<title>" & GetNewTitle(strScriptName) & "</title>" & strLE
'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<meta name=""dcterms.rights"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
'## "<link href=""normalize.css"" rel=""stylesheet"" media=""all"">" & strLE & _
Response.Write _
	"<meta name=""description"" content="""">" & strLE & _
	"<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & strLE & _
	"<link href=""css/normalize-legacy.css"" rel=""stylesheet"" media=""all"">" & strLE & _
	"<link href=""css/snitz.css"" rel=""stylesheet"" media=""all"">" & strLE & _
	"</head>" & strLE & strLE & _
	"<body" & strTmpPageBGImageURL & " id=""top"">" & strLE & strLE & _
	"<header role=""banner"">" & strLE & _
	"<div id=""header"">" & strLE & _
	"<div class=""hdlogo""><a href=""default.asp"">" & getCurrentIcon(strTitleImage & "||",strForumTitle,"") & "</a></div>" & strLE & _
	"<div class=""hdcontrols"">" & strLE
Response.Write "<b>" & chkString(strForumTitle,"pagetitle") & "</b><br>" & strLE & _
	"<div class=""hdnav"">" & strLE & _
	"<nav role=""navigation"">" & strLE
Call sForumNavigation()
Response.Write strLE & "</nav>" & strLE & _
	"</div>" & strLE & _
	"<!-- /hdnav -->" & strLE

if (mlev = 0) then
	Response.Write "<form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form1"" name=""form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""login"">" & strLE
else
	Response.Write "<form action=""" & Request.ServerVariables("URL") & """ method=""post"" id=""form2"" name=""form2"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""logout"">" & strLE
end if

Call ChkLoggingAction()

if (mlev = 0) then
	if not(Instr(Request.ServerVariables("Path_Info"), "register.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "pop_profile.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "search.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "login.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "password.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "faq.asp") > 0) and _
	not(Instr(Request.ServerVariables("Path_Info"), "post.asp") > 0) then
		Response.Write "<table class=""login"">" & strLE & _
			"<tr class=""lstatus"">" & strLE
		if (strAuthType = "db") then
			Response.Write "<td><b>Username:</b><br><input type=""text"" name=""Name"" size=""10"" maxLength=""25"" value=""""></td>" & strLE & _
				"<td><b>Password:</b><br><input type=""password"" name=""Password"" size=""10"" maxLength=""25"" value=""""></td>" & strLE & _
				"<td>"
			if strGfxButtons = "1" then
				Response.Write "<input type=""image"" src=""" & strImageUrl & "button_login.gif"" alt=""Login"" id=""submit1"" name=""Login"">"
			else
				Response.Write "<input type=""submit"" value=""Login"" id=""submit1"" name=""submit1"">"
			end if
			Response.Write "</td>" & strLE & _
				"</tr>" & strLE & _
				"<tr class=""loptions"">" & strLE & _
				"<td colspan=""3"">" & strLE & _
				"<input type=""checkbox"" name=""SavePassWord"" value=""true"" checked> <b>Save Password</b></td>" & strLE
		else
			if (strAuthType = "nt") then
				Response.Write "<td class=""lmessage""><p>Please <a href=""register.asp"">register</a> to post in these Forums</p></td>" & strLE
			end if
		end if
		Response.Write "</tr>" & strLE
		if (lcase(strEmail) = "1") then
			Response.Write "<tr class=""loptions"">" & strLE & _
				"<td colspan=""3"">" & strLE & _
				"<a href=""password.asp"">Forgot your "
			if strAuthType = "nt" then Response.Write "Admin "
			Response.Write "Password?</a>" & strLE
			if (lcase(strNoCookies) = "1") then
				Response.Write " | <a href=""admin_home.asp"">Admin Options</a>" & strLE
			end if
			Response.Write "</td>" & strLE & _
				"</tr>" & strLE
		end if
		Response.Write "</table>" & strLE & _
			"<br><br>" & srtLE
	end if
else
	Response.Write "<table class=""logout"">" & strLE & _
		"<tr class=""lstatus"">" & strLE & _
		"<td>You are logged on as<br>"
	if strAuthType="nt" then
		Response.Write "<b>" & Session(strCookieURL & "username") & "&nbsp;(" & Session(strCookieURL & "userid") & ")</b></td>" & strLE & _
			"<td>&nbsp;"
	else
		if strAuthType = "db" then
			Response.Write "<b>" & profileLink(ChkString(strDBNTUserName, "display"),MemberID) & "</b></td>" & strLE & _
				"<td>"
			if strGfxButtons = "1" then
				Response.Write "<input src=""" & strImageUrl & "button_logout.gif"" alt=""Logout"" type=""image"" id=""submit1"" name=""Logout"">"
			else
				Response.Write "<input type=""submit"" value=""Logout"" id=""submit1"" name=""submit1"">"
			end if
		end if
	end if
	Response.Write "</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE
	if (mlev = 4) or (lcase(strNoCookies) = "1") then
		Response.Write "<div class=""loptions"">" & strLE & _
			"<a href=""admin_home.asp"">Admin Options</a>"
		if mLev = 4 and (strEmailVal = "1" _
		and strRestrictReg = "1" _
		and strEmail = "1" _
		and User_Count > 0) then
			Response.Write "&nbsp;|&nbsp;<a href=""admin_accounts_pending.asp"">(" & User_Count & ") Member(s) awaiting approval</a>"
		end if
		Response.Write "</div>" & strLE
	end if
	Response.Write "<br><br>" & strLE
end if
Response.Write "</form>" & strLE & _
	"</div>" & strLE & _
	"<!-- /hdcontrols -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /header -->" & strLE & _
	"</header>" & strLE & strLE //& _
	//"<main role=""main"">" & strLE
'########### GROUP Categories ########### %>
<!--#INCLUDE FILE="inc_groupjump_to.asp" -->
<% '######## GROUP Categories ##############

sub sForumNavigation()
	' DEM --> Added code to show the subscription line
	if strSubscription > 0 and strEmail = "1" then
		if mlev > 0 then
			strSql = "SELECT COUNT(*) AS MySubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
			strSql = strSql & " WHERE MEMBER_ID = " & MemberID
			set rsCount = my_Conn.Execute (strSql)
			if rsCount.BOF or rsCount.EOF then
				' No Subscriptions found, do nothing
				MySubCount = 0
				rsCount.Close
				set rsCount = nothing
			else
				MySubCount = rsCount("MySubCount")
				rsCount.Close
				set rsCount = nothing
			end if
			if mLev = 4 then
				strSql = "SELECT COUNT(*) AS SubCount FROM " & strTablePrefix & "SUBSCRIPTIONS"
				set rsCount = my_Conn.Execute (strSql)
				if rsCount.BOF or rsCount.EOF then
					' No Subscriptions found, do nothing
					SubCount = 0
					rsCount.Close
					set rsCount = nothing
				else
					SubCount = rsCount("SubCount")
					rsCount.Close
					set rsCount = nothing
				end if
			end if
		else
			SubCount = 0
			MySubCount = 0
		end if
	else
		SubCount = 0
		MySubCount = 0
	end if
	Response.Write "<a href=""" & strHomeURL & """>Home</a>" & strLE
	if strUseExtendedProfile then
		Response.Write " | <a href=""pop_profile.asp?mode=Edit"">Profile</a>" & strLE
	else
		Response.Write " | <a href=""javascript:openWindow3('pop_profile.asp?mode=Edit')"">Profile</a>" & strLE
	end if
	if strAutoLogon <> "1" then
		if strProhibitNewMembers <> "1" then
			Response.Write " | <a href=""register.asp"">Register</a>" & strLE
		end if
	end if
	Response.Write " | <a href=""active.asp"">Active Topics</a>" & strLE
	' DEM --> Start of code added to show subscriptions if they exist
	if (strSubscription > 0) then
		if mlev = 4 and SubCount > 0 then
			Response.Write " | <a href=""subscription_list.asp?MODE=all"">All Subscriptions</a>" & strLE
		end if
		if MySubCount > 0 then
			Response.Write " | <a href=""subscription_list.asp"">My Subscriptions</a>" & strLE
		end if
	end if
	' DEM --> End of Code added to show subscriptions if they exist
	Response.Write " | <a href=""members.asp"">Members</a>" & strLE & _
		" | <a href=""search.asp"
	if Request.QueryString("FORUM_ID") <> "" then
		Response.Write "?FORUM_ID=" & cLng(Request.QueryString("FORUM_ID"))
	end if
	Response.Write """>Search</a>" & strLE & _
		" | <a href=""faq.asp"">FAQ</a>"
end sub

if strGroupCategories = "1" then
	if Session(strCookieURL & "GROUP_NAME") = "" then
		GROUPNAME = " Default Groups "
	else
		GROUPNAME = Session(strCookieURL & "GROUP_NAME")
	end if
	'Forum_SQL - Get Groups
	strSql = "SELECT GROUP_ID, GROUP_CATID "
	strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
	strSql = strSql & " WHERE GROUP_ID = " & Group
	set rsgroups = Server.CreateObject("ADODB.Recordset")
	rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rsgroups.EOF then
		recGroupCatCount = ""
	else
		allGroupCatData = rsgroups.GetRows(adGetRowsRest)
		recGroupCatCount = UBound(allGroupCatData, 2)
	end if
	rsgroups.Close
	set rsgroups = nothing
end if

sub ChkLoggingAction()
	select case Request.Form("Method_Type")
		case "login"
			Response.Write "</div>" & strLE & _
				"<!-- /hcontrols -->" & strLE & _
				"</div>" & strLE & _
				"<!-- /header -->" & strLE & _
				"<div class=""lmessage"">" & strLE
			if strLoginStatus = 0 then
				Response.Write "<p>Your username and/or password were incorrect</p>" & strLE & _
					"<p>Please either try again or register for an account</p>" & strLE
			else
				Response.Write "<p>You logged on successfully!</p>" & strLE & _
					"<p>Thank you for your participation</p>" & strLE
			end if
			Response.Write "<meta http-equiv=""Refresh"" content=""2; URL=" & strReferer & """>" & strLE & _
				"<a href=""" & strReferer & """>Back To Forum</a>" & strLE & _
				"</div>" & strLE & _
				"<!-- /lmessage -->" & strLE & _
				"</header>" & strLE & strLE & _
				"<main role=""main"">" & strLE & _
				"<table>" & strLE & _
				"<tr>" & strLE & _
				"<td>" & strLE
			Call WriteFooter
			Response.End
		case "logout"
			Response.Write "</div><!-- .hcontrols -->" & strLE & _
				"</div>" & strLE & _
				"<!-- /header -->" & strLE & _
				"<div class=""lmessage"">" & strLE & _
				"<p>You logged out successfully!</p>" & strLE & _
				"<p>Thank you for your participation.</p>" & strLE & _
				"<meta http-equiv=""Refresh"" content=""2; URL=default.asp"">" & strLE & _
				"<a href=""default.asp"">Back To Forum</a>" & strLE & _
				"</div>" & strLE & _
				"<!-- /lmessage -->" & strLE & _
				"</header>" & strLE & strLE & _
				"<main role=""main"">" & strLE & _
				"<table>" & strLE & _
				"<tr>" & strLE & _
				"<td>" & strLE
			Call WriteFooter
			Response.End
	end select
end sub
%>
