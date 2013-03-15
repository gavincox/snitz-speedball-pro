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
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Feature&nbsp;Configuration<br>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE

if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strIMGInPosts") = "1" and Request.Form("strAllowForumCode") = "0" then
		Err_Msg = Err_Msg & "<li>Forum Code Must be Enabled in order to Enable Images</li>"
	end if
	if Request.Form("strAllowHTML") = "1" and Request.Form("strAllowForumCode") = "1" then
		Err_Msg = Err_Msg & "<li>HTML and ForumCode cannot both be On at the same time</li>"
	end if
	if Request.Form("intHotTopicNum") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter a Hot Topic Number</li>"
	elseif IsNumeric(Request.Form("intHotTopicNum")) = False then
		Err_Msg = Err_Msg & "<li>Hot Topic Number must be a number</li>"
	elseif cLng(Request.Form("intHotTopicNum")) = 0 then
		Err_Msg = Err_Msg & "<li>Hot Topic Number cannot be 0</li>"
	end if
	if left(Request.Form("intHotTopicNum"), 1) = "-" then
		Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number</li>"
	end if
	if left(Request.Form("intHotTopicNum"), 1) = "+" then
		Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number without the <b>+</b></li>"
	end if
	if Request.Form("strPageSize") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter the number of Items per Page</li>"
	elseif IsNumeric(Request.Form("strPageSize")) = False then
		Err_Msg = Err_Msg & "<li>Items per Page must be a number</li>"
	elseif cLng(Request.Form("strPageSize")) = 0 then
		Err_Msg = Err_Msg & "<li>Items per Page cannot be 0</li>"
	end if
	if Request.Form("strPageNumberSize") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter the number of Pages per Row</li>"
	elseif IsNumeric(Request.Form("strPageNumberSize")) = False then
		Err_Msg = Err_Msg & "<li>Pages per Row must be a number</li>"
	elseif cLng(Request.Form("strPageNumberSize")) = 0 then
		Err_Msg = Err_Msg & "<li>Pages per Row cannot be 0</li>"
	end if

	if (strShowTimer = "1" or Request.Form("strShowTimer") = "1") and Request.Form("strShowTimer") <> "0" then
		if trim(Request.Form("strTimerPhrase")) = "" then
			Err_Msg = Err_Msg & "<li>You Must Enter a Phrase for the Timer</li>"
		end if
		if Instr(Request.Form("strTimerPhrase"), "[TIMER]") = "0" then
			Err_Msg = Err_Msg & "<li>Your Timer Phrase must contain the [TIMER] placeholder</li>"
		end if
	end if
	if strModeration = "1" and Request.Form("strModeration") = "0" then
        	if CheckForUnmoderatedPosts("BOARD", 0, 0, 0) > 0 then
			Err_Msg = Err_Msg & "<li>Please Approve or Delete all UnModerated/Held posts before turning Moderation off.</li>"
		end if
	end if

	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next

		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write "<p class=""c""><span class=""dff hfs"">Configuration Posted!</span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & strLE & _
			"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""admin_home.asp"">Back To Admin Home</span></a></p>" & strLE
	else
		Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
			"<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
	end if
else
	Response.Write "<form action=""admin_config_features.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Feature Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Security Settings</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Secure Admin Mode</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strSecureAdmin"" value=""1""" & chkRadio(strSecureAdmin,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSecureAdmin"" value=""0""" & chkRadio(strSecureAdmin,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#secureadminmode')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Non-Cookie Mode</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strNoCookies"" value=""1""" & chkRadio(strNoCookies,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strNoCookies"" value=""0""" & chkRadio(strNoCookies,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#allownoncookies')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>General Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>IP Logging</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strIPLogging"" value=""1""" & chkRadio(strIPLogging,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strIPLogging"" value=""0""" & chkRadio(strIPLogging,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#IPLogging')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Flood Control</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strFloodCheck"" value=""1""" & chkRadio(strFloodCheck,0,false) & "> On" & strLE & _
		"<select name=""strFloodCheckTime"">" & strLE & _
		"<option value=""-30""" & chkSelect(strFloodCheckTime,-30) & ">30 seconds</option>" & strLE & _
		"<option value=""-60""" & chkSelect(strFloodCheckTime,-60) & ">60 seconds</option>" & strLE & _
		"<option value=""-90""" & chkSelect(strFloodCheckTime,-90) & ">90 seconds</option>" & strLE & _
		"<option value=""-120""" & chkSelect(strFloodCheckTime,-120) & ">120 seconds</option>" & strLE & _
		"</select>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFloodCheck"" value=""0""" & chkRadio(strFloodCheck,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#FloodCheck')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Private Forums</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strPrivateForums"" value=""1""" & chkRadio(strPrivateForums,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strPrivateForums"" value=""0""" & chkRadio(strPrivateForums,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#privateforums')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Group Categories</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strGroupCategories"" value=""1""" & chkRadio(strGroupCategories,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strGroupCategories"" value=""0""" & chkRadio(strGroupCategories,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#groupcategories')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Highest level of Subscription</b>&nbsp;</td>" & strLE & _
		"<td><select name=""strSubscription"">" & strLE & _
		"<option value=""0""" & chkSelect(strSubscription,0) & ">No Subscriptions Allowed</option>" & strLE & _
		"<option value=""1""" & chkSelect(strSubscription,1) & ">Subscribe to Whole Board</option>" & strLE & _
		"<option value=""2""" & chkSelect(strSubscription,2) & ">Subscribe by Category</option>" & strLE & _
		"<option value=""3""" & chkSelect(strSubscription,3) & ">Subscribe by Forum</option>" & strLE & _
		"<option value=""4""" & chkSelect(strSubscription,4) & ">Subscribe by Topic</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#Subscription')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Bad Word Filter</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strBadWordFilter"" value=""1""" & chkRadio(strBadWordFilter,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strBadWordFilter"" value=""0""" & chkRadio(strBadWordFilter,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#badwordfilter')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Moderation Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow Topic Moderation</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strModeration"" value=""1""" & chkRadio(strModeration,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strModeration"" value=""0""" & chkRadio(strModeration,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#Moderation')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Moderators</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowModerators"" value=""1""" & chkRadio(strShowModerators,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowModerators"" value=""0""" & chkRadio(strShowModerators,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowModerator')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Restrict Moderators to&nbsp;&nbsp;<br> moving their own topics</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strMoveTopicMode"" value=""1""" & chkRadio(strMoveTopicMode,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMoveTopicMode"" value=""0""" & chkRadio(strMoveTopicMode,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#MoveTopicMode')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>AutoEmail author&nbsp;&nbsp;<br>when moving topics</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strMoveNotify"" value=""1""" & chkRadio(strMoveNotify,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMoveNotify"" value=""0""" & chkRadio(strMoveNotify,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#MoveNotify')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Forum Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Archive Functions</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strArchiveState"" value=""1""" & chkRadio(strArchiveState,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strArchiveState"" value=""0""" & chkRadio(strArchiveState,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ArchiveState')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Detailed Statistics</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowStatistics"" value=""1""" & chkRadio(strShowStatistics,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowStatistics"" value=""0""" & chkRadio(strShowStatistics,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#stats')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Jump To Last Post Link</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strJumpLastPost"" value=""1""" & chkRadio(strJumpLastPost,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strJumpLastPost"" value=""0""" & chkRadio(strJumpLastPost,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#JumpLastPost')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Quick Paging</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowPaging"" value=""1""" & chkRadio(strShowPaging,0,false) & ">&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowPaging"" value=""0""" & chkRadio(strShowPaging,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowPaging')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Pagenumbers per row</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPageNumberSize"" size=""5"" maxLength=""3"" value=""" & chkExistElse(strPageNumbersize,10) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#pagenumbersize')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Topic Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow Sticky Topics</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strStickyTopic"" value=""1""" & chkRadio(strStickyTopic,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strStickyTopic"" value=""0""" & chkRadio(strStickyTopic,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#StickyTopic')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Edited By on Date</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strEditedByDate"" value=""1""" & chkRadio(strEditedByDate,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strEditedByDate"" value=""0""" & chkRadio(strEditedByDate,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#editedbydate')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Prev / Next Topic</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowTopicNav"" value=""1""" & chkRadio(strShowTopicNav,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowTopicNav"" value=""0""" & chkRadio(strShowTopicNav,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowTopicNav')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Send Topic to a Friend Link</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowSendToFriend"" value=""1""" & chkRadio(strShowSendToFriend,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowSendToFriend"" value=""0""" & chkRadio(strShowSendToFriend,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowSendToFriend')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Printer Friendly Link</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowPrinterFriendly"" value=""1""" & chkRadio(strShowPrinterFriendly,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowPrinterFriendly"" value=""0""" & chkRadio(strShowPrinterFriendly,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowPrinterFriendly')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Hot Topics</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strHotTopic"" value=""1""" & chkRadio(strHotTopic,0,false) & "> On" & strLE & _
		"<input type=""text"" name=""intHotTopicNum"" size=""5"" maxLength=""3"" value=""" & chkExistElse(intHotTopicNum,20) & """>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strHotTopic"" value=""0""" & chkRadio(strHotTopic,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#hottopics')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Items per page</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPageSize"" size=""5"" maxLength=""3"" value=""" & chkExistElse(strPageSize,15) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#pagesize')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Posting Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow HTML</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strAllowHTML"" value=""1""" & chkRadio(strAllowHTML,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAllowHTML"" value=""0""" & chkRadio(strAllowHTML,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#AllowHTML')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow Forum Code</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strAllowForumCode"" value=""1""" & chkRadio(strAllowForumCode,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAllowForumCode"" value=""0""" & chkRadio(strAllowForumCode,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#AllowForumCode')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Images in Posts</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strIMGInPosts"" value=""1""" & chkRadio(strIMGInPosts,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strIMGInPosts"" value=""0""" & chkRadio(strIMGInPosts,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#imginposts')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Icons</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strIcons"" value=""1""" & chkRadio(strIcons,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strIcons"" value=""0""" & chkRadio(strIcons,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#icons')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow Signatures</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strSignatures"" value=""1""" & chkRadio(strSignatures,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSignatures"" value=""0""" & chkRadio(strSignatures,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#signatures')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Allow Dynamic Signatures</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strDSignatures"" value=""1""" & chkRadio(strDSignatures,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strDSignatures"" value=""0""" & chkRadio(strDSignatures,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#dsignatures')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Format Buttons</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowFormatButtons"" value=""1""" & chkRadio(strShowFormatButtons,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowFormatButtons"" value=""0""" & chkRadio(strShowFormatButtons,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowFormatButtons')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Smilies Table</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowSmiliesTable"" value=""1""" & chkRadio(strShowSmiliesTable,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowSmiliesTable"" value=""0""" & chkRadio(strShowSmiliesTable,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowSmiliesTable')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Quick Reply</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowQuickReply"" value=""1""" & chkRadio(strShowQuickReply,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowQuickReply"" value=""0""" & chkRadio(strShowQuickReply,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#ShowQuickReply')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Misc Features</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Timer</b>&nbsp;</td>" & strLE & _
		"<td><input type=""radio"" class=""radio"" name=""strShowTimer"" value=""1""" & chkRadio(strShowTimer,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowTimer"" value=""0""" & chkRadio(strShowTimer,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#timer')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Timer Phrase</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strTimerPhrase"" size=""45"" maxLength=""50"" value=""" & chkExistElse(strTimerPhrase,"This page was generated in [TIMER] seconds.") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=features#timerphrase')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE
end if
WriteFooter
Response.end
%>
