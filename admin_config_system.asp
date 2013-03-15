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
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Main&nbsp;Forum&nbsp;Configuration" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strTitleImage") = "" then Err_Msg = Err_Msg & "<li>You Must Enter the Address of a Title Image</li>"
	if Request.Form("strHomeURL") = "" then Err_Msg = Err_Msg & "<li>You Must Enter the URL of your HomePage (either relative or full)</li>"
	if Request.Form("strForumURL") = "" then Err_Msg = Err_Msg & "<li>You Must Enter the Fully Qualified URL of your Forum</li>"
	if (left(lcase(Request.Form("strForumURL")), 7) <> "http://" _
	and left(lcase(Request.Form("strForumURL")), 8) <> "https://") _
	and Request.Form("strHomeURL") <> "" then Err_Msg = Err_Msg & "<li>You Must prefix the Forum URL with <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
	if (right(lcase(Request.Form("strForumURL")), 1) <> "/") then Err_Msg = Err_Msg & "<li>You Must end the Forum URL with <b>/</b></li>"
	if trim(Request.Form("strImageURL")) <> "" then
		if (right(lcase(Request.Form("strImageURL")), 1) <> "/") then Err_Msg = Err_Msg & "<li>You Must end the Images Location with <b>/</b></li>"
	end if
	if Request.Form("strAuthType") <> strAuthType and strAuthType = "db" then
		if not(mLev = 4 and MemberID = intAdminMemberID) then
			Err_Msg = Err_Msg & "<li>Only the Admin user can change the Authentication type of the board</li>"
		else
			call NTauthenticate()
			if Session(strCookieURL & "userid") = "" then
				Err_Msg = Err_Msg & "<li>You have to enable non-Anonymous access for the forum on the server first</li>"
			else
				strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET " & strMemberTablePrefix & "MEMBERS.M_USERNAME = '" & Session(strCookieURL & "userid") & "'"
				strSql = strSql & " WHERE " & strMemberTablePrefix & "MEMBERS.M_NAME = '" & Request.Cookies(strUniqueID & "User")("Name") & "'"
				my_Conn.Execute(strSql),,adCmdText + adExecuteNoRecords
				call NTauthenticate()
				call NTUser()
			end if
		end if
	end if
	if (Request.Form("strAuthType") <> strAuthType) and strAuthType = "nt" then
		if not(mLev = 4 and MemberID = intAdminMemberID) then
			Err_Msg = Err_Msg & "<li>Only the Admin user can change the Authentication type of the board</li>"
		else
			Session(strCookieURL & "Approval") = ""
		end if
	end if
	if Err_Msg = "" then
		'## Forum_SQL
		for each key in Request.Form
			if left(key,3) = "str" then strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
		next
		Application(strCookieURL & "ConfigLoaded") = ""
		Response.Write "<p class=""c""><span class=""dff hfs"">Configuration Posted!</span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & strLE & _
			"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
			"<p class=""c""><a href=""admin_home.asp"">Back To Admin Home</a></p>" & strLE
	else
		Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
			"<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></p>" & strLE
	end if
else
	Response.Write "<form action=""admin_config_system.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Main Forum Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Forum's Title</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumTitle"" size=""30"" value=""" & chkExist(chkString(strForumTitle,"edit")) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#forumtitle')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Forum's Copyright</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strCopyright"" size=""30"" value=""" & chkExistElse(chkString(strCopyright,"edit"),"2000-2002 Snitz Communications") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#copyright')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Title Image Location</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strTitleImage"" size=""30"" value=""" & chkExistElse(strTitleImage,"logo_snitz_forums_2000.gif") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#titleimage')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Home URL</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strHomeURL"" size=""30"" value=""" & chkExistElse(strHomeURL,"../") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#homeurl')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Forum URL</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumURL"" size=""30"" value=""" & chkExistElse(strForumURL,"./") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#forumurl')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Images Location</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strImageURL"" size=""30"" value=""" & chkExist(strImageURL) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#imagelocation')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Version info</b>&nbsp;</td>" & strLE & _
		"<td>"
	if strVersion <> "" then Response.Write("[<i>"& strVersion & "</i>]") else Response.Write("<b>[Couldn't read version info..]</b>")
	Response.Write "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Authorization Type</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAuthType"" value=""db""" & chkRadio(strAuthType,"db",true) & "> DB" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAuthType"" value=""nt""" & chkRadio(strAuthType,"nt",true) & "> NT" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#AuthType')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Set Cookie To</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSetCookieToForum"" value=""1""" & chkRadio(strSetCookieToForum,1,true) & "> Forum" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSetCookieToForum"" value=""0""" & chkRadio(strSetCookieToForum,0,true) & "> Website" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#SetCookieToForum')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Use Graphics as Buttons</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strGfxButtons"" value=""1""" & chkRadio(strGfxButtons,1,true) & "> On" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strGfxButtons"" value=""0""" & chkRadio(strGfxButtons,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#GfxButtons')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Use Graphic for ""Powered By"" link</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowImagePoweredBy"" value=""1""" & chkRadio(strShowImagePoweredBy,1,true) & "> On" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strShowImagePoweredBy"" value=""0""" & chkRadio(strShowImagePoweredBy,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#PoweredBy')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Prohibit New Members</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strProhibitNewMembers"" value=""1""" & chkRadio(strProhibitNewMembers,1,true) & "> On" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strProhibitNewMembers"" value=""0""" & chkRadio(strProhibitNewMembers,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#ProhibitNewMembers')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Require Registration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRequireReg"" value=""1""" & chkRadio(strRequireReg,1,true) & "> On" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRequireReg"" value=""0""" & chkRadio(strRequireReg,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#RequireReg')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>UserName Filter</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strUserNameFilter"" value=""1""" & chkRadio(strUserNameFilter,1,true) & "> On" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strUserNameFilter"" value=""0""" & chkRadio(strUserNameFilter,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=system#UserNameFilter')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE
end if
Call WriteFooter
Response.End
%>
