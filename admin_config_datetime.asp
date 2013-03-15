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
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Server&nbsp;Date/Time&nbsp;Configuration" & strLE & _
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
	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
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
	Response.Write "<form action=""admin_config_datetime.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Server Date/Time Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Time Display</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strTimeType"" value=""12""" & chkRadio(strTimeType,12,true) & "> 12hr&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strTimeType"" value=""24""" & chkRadio(strTimeType,24,true) & "> 24hr" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#timetype')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Time Adjustment</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strTimeAdjust"">" & strLE
	for iTimeAdjust = -24 to 24
		Response.Write "<option value=""" & iTimeAdjust & """" & chkSelect(strTimeAdjust,iTimeAdjust) & ">" & iTimeAdjust & "</option>" & strLE
	next
	Response.Write "</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#TimeAdjust')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Current Forum Date/Time</b>&nbsp;</td>" & strLE & _
		"<td>&nbsp;" & ChkDate(datetostr(strForumTimeAdjust),"",true) & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Date Display</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strDateType"">" & strLE & _
		"<option value=""mdy""" & chkSelect(strDateType,"mdy") & ">12/31/2000 (US short)</option>" & strLE & _
		"<option value=""dmy""" & chkSelect(strDateType,"dmy") & ">31/12/2000 (UK short)</option>" & strLE & _
		"<option value=""ymd""" & chkSelect(strDateType,"ymd") & ">2000/12/31 (Other short)</option>" & strLE & _
		"<option value=""ydm""" & chkSelect(strDateType,"ydm") & ">2000/31/12 (Other short)</option>" & strLE & _
		"<option value=""mmdy""" & chkSelect(strDateType,"mmdy") & ">Dec 31 2000 (US med)</option>" & strLE & _
		"<option value=""dmmy""" & chkSelect(strDateType,"dmmy") & ">31 Dec 2000 (UK med)</option>" & strLE & _
		"<option value=""ymmd""" & chkSelect(strDateType,"ymmd") & ">2000 Dec 31 (Other med)</option>" & strLE & _
		"<option value=""ydmm""" & chkSelect(strDateType,"ydmm") & ">2000 31 Dec (Other med)</option>" & strLE & _
		"<option value=""mmmdy""" & chkSelect(strDateType,"mmmdy") & ">December 31 2000 (US long)</option>" & strLE & _
		"<option value=""dmmmy""" & chkSelect(strDateType,"dmmmy") & ">31 December 2000 (UK long)</option>" & strLE & _
		"<option value=""ymmmd""" & chkSelect(strDateType,"ymmmd") & ">2000 December 31 (Other long)</option>" & strLE & _
		"<option value=""ydmmm""" & chkSelect(strDateType,"ydmmm") & ">2000 31 December (Other long)</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=datetime#datetype')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE
end if
Call WriteFooter
Response.End
%>
