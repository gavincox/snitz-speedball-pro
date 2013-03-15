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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" Then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Member&nbsp;Details&nbsp;Configuration" & strLE & _
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
	if Request.Form("strAge") = "1" and Request.Form("strAgeDOB") = "1" then Err_Msg = Err_Msg & "<li>Age and Birth Date cannot both be On at the same time</li>"
	intAge = ChkString(trim(Request.Form("strMinAge")), "SQLString")
	if len(intAge) = 0 then intAge = 0
	if not isNumeric(intAge) then Err_Msg = Err_Msg & "<li>Minimum Age must be a numerical value.</li>"
	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" then strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
		next
		Application(strCookieURL & "ConfigLoaded") = ""
		Response.Write "<p class=""c""><span class=""dff hfs"">Configuration Posted!</span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & strLE & _
			"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
			"<p class=""c""><a href=""admin_home.asp"">Back To Admin Home</span></a></p>" & strLE
	else
		Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
			"<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
	end if
else
	Response.Write "<form action=""admin_config_members.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Member Details Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Full Name</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFullName"" value=""1""" & chkRadio(strFullName,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFullName"" value=""0""" & chkRadio(strFullName,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#FullName')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqFullName"" value=""1""" & chkRadio(strReqFullName,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqFullName"" value=""0""" & chkRadio(strReqFullName,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Picture</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strPicture"" value=""1""" & chkRadio(strPicture,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strPicture"" value=""0""" & chkRadio(strPicture,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Picture')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqPicture"" value=""1""" & chkRadio(strReqPicture,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqPicture"" value=""0""" & chkRadio(strReqPicture,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Recent Topics</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRecentTopics"" value=""1""" & chkRadio(strRecentTopics,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRecentTopics"" value=""0""" & chkRadio(strRecentTopics,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#RecentTopics')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Gender (male/female)</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSex"" value=""1""" & chkRadio(strSex,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strSex"" value=""0""" & chkRadio(strSex,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Sex')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqSex"" value=""1""" & chkRadio(strReqSex,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqSex"" value=""0""" & chkRadio(strReqSex,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Age</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAge"" value=""1""" & chkRadio(strAge,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAge"" value=""0""" & chkRadio(strAge,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Age')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAge"" value=""1""" & chkRadio(strReqAge,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAge"" value=""0""" & chkRadio(strReqAge,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Birth Date</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAgeDOB"" value=""1""" & chkRadio(strAgeDOB,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAgeDOB"" value=""0""" & chkRadio(strAgeDOB,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#AgeDOB')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAgeDOB"" value=""1""" & chkRadio(strReqAgeDOB,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAgeDOB"" value=""0""" & chkRadio(strReqAgeDOB,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Minimum Age</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""putc c"" colspan=""2"">" & strLE
	intYoungest = 0
	if strAge = "1" then
		set rs = my_Conn.execute(TopSQL("SELECT M_AGE FROM " & strMemberTablePrefix & "MEMBERS WHERE M_AGE <> '' AND M_STATUS <> 0 ORDER BY M_AGE ASC", 1))
		if not rs.eof then intYoungest = cInt(rs("M_AGE"))
		rs.close
		set rs = nothing
	elseif strAgeDOB = "1" then
		set rs = my_Conn.execute(TopSQL("SELECT M_DOB FROM " & strMemberTablePrefix & "MEMBERS WHERE M_DOB <> '' AND M_STATUS <> 0 ORDER BY M_DOB DESC", 1))
		if not rs.eof then intYoungest = cInt(DisplayUsersAge(DOBToDate(rs("M_DOB"))))
		rs.close
		set rs = nothing
	end if
	if intYoungest > 0 then Response.Write "<span class=""dff ffs hlfc""><b>Youngest member is " & intYoungest & "&nbsp;</b></span><br>" & strLE
	Response.Write "<input type=""text"" name=""strMinAge"" value="""&strMinAge&""" maxlength=""2"" size=""16""> " & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#MinAge')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>City</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strCity"" value=""1""" & chkRadio(strCity,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strCity"" value=""0""" & chkRadio(strCity,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#City')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqCity"" value=""1""" & chkRadio(strReqCity,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqCity"" value=""0""" & chkRadio(strReqCity,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>State</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strState"" value=""1""" & chkRadio(strState,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strState"" value=""0""" & chkRadio(strState,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#State')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqState"" value=""1""" & chkRadio(strReqState,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqState"" value=""0""" & chkRadio(strReqState,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Country</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strCountry"" value=""1""" & chkRadio(strCountry,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strCountry"" value=""0""" & chkRadio(strCountry,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Country')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqCountry"" value=""1""" & chkRadio(strReqCountry,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqCountry"" value=""0""" & chkRadio(strReqCountry,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>AIM</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAIM"" value=""1""" & chkRadio(strAIM,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strAIM"" value=""0""" & chkRadio(strAIM,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#aim')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAIM"" value=""1""" & chkRadio(strReqAIM,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqAIM"" value=""0""" & chkRadio(strReqAIM,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>ICQ</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strICQ"" value=""1""" & chkRadio(strICQ,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strICQ"" value=""0""" & chkRadio(strICQ,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#icq')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqICQ"" value=""1""" & chkRadio(strReqICQ,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqICQ"" value=""0""" & chkRadio(strReqICQ,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>MSN</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMSN"" value=""1""" & chkRadio(strMSN,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMSN"" value=""0""" & chkRadio(strMSN,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#msn')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqMSN"" value=""1""" & chkRadio(strReqMSN,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqMSN"" value=""0""" & chkRadio(strReqMSN,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Yahoo!</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strYAHOO"" value=""1""" & chkRadio(strYAHOO,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strYAHOO"" value=""0""" & chkRadio(strYAHOO,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#yahoo')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqYAHOO"" value=""1""" & chkRadio(strReqYAHOO,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqYAHOO"" value=""0""" & chkRadio(strReqYAHOO,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Occupation</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strOccupation"" value=""1""" & chkRadio(strOccupation,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strOccupation"" value=""0""" & chkRadio(strOccupation,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Occupation')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqOccupation"" value=""1""" & chkRadio(strReqOccupation,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqOccupation"" value=""0""" & chkRadio(strReqOccupation,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Homepages</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strHomepage"" value=""1""" & chkRadio(strHomepage,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strHomepage"" value=""0""" & chkRadio(strHomepage,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Homepages')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqHomepage"" value=""1""" & chkRadio(strReqHomepage,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqHomepage"" value=""0""" & chkRadio(strReqHomepage,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Favorite Links</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFavLinks"" value=""1""" & chkRadio(strFavLinks,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFavLinks"" value=""0""" & chkRadio(strFavLinks,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#FavLinks')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqFavLinks"" value=""1""" & chkRadio(strReqFavLinks,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqFavLinks"" value=""0""" & chkRadio(strReqFavLinks,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Marital Status</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMarStatus"" value=""1""" & chkRadio(strMarStatus,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strMarStatus"" value=""0""" & chkRadio(strMarStatus,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#MStatus')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqMarStatus"" value=""1""" & chkRadio(strReqMarStatus,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqMarStatus"" value=""0""" & chkRadio(strReqMarStatus,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Bio</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strBio"" value=""1""" & chkRadio(strBio,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strBio"" value=""0""" & chkRadio(strBio,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Bio')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqBio"" value=""1""" & chkRadio(strReqBio,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqBio"" value=""0""" & chkRadio(strReqBio,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Hobbies</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strHobbies"" value=""1""" & chkRadio(strHobbies,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strHobbies"" value=""0""" & chkRadio(strHobbies,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#hobbies')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqHobbies"" value=""1""" & chkRadio(strReqHobbies,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqHobbies"" value=""0""" & chkRadio(strReqHobbies,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Latest News</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strLNews"" value=""1""" & chkRadio(strLNews,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strLNews"" value=""0""" & chkRadio(strLNews,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#LNews')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqLNews"" value=""1""" & chkRadio(strReqLNews,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqLNews"" value=""0""" & chkRadio(strReqLNews,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Quote</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Enabled</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strQuote"" value=""1""" & chkRadio(strQuote,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strQuote"" value=""0""" & chkRadio(strQuote,0,true) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=members#Quote')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Required</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqQuote"" value=""1""" & chkRadio(strReqQuote,0,false) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strReqQuote"" value=""0""" & chkRadio(strReqQuote,0,true) & "> Off" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE
end if
Call WriteFooter
%>
