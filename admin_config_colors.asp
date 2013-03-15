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
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Font/Table&nbsp;Color&nbsp;Code&nbsp;Configuration" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strTopicWidthLeft") = "" then Err_Msg = Err_Msg & "<li>You Must enter a value for the Topic Left Column Width</li>"
	if Request.Form("strTopicWidthRight") = "" then Err_Msg = Err_Msg & "<li>You Must enter a value for the Topic Right Column Width</li>"
	if Err_Msg = "" then
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
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
	Response.Write "<form action=""admin_config_colors.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Font/Table Color Code Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Font Face Type</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strDefaultFontFace"" size=""25"" maxLength=""30"" value=""" & chkExist(strDefaultFontFace) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontfacetype')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Default Font Size</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strDefaultFontSize"">" & strLE & _
		"<option value=""""" & chkSelect(strDefaultFontSize,"") & ">None (blank)</option>" & strLE & _
		"<option value=""1""" & chkSelect(strDefaultFontSize,1) & ">1 (8 pt)</option>" & strLE & _
		"<option value=""2""" & chkSelect(strDefaultFontSize,2) & ">2 (10 pt)</option>" & strLE & _
		"<option value=""3""" & chkSelect(strDefaultFontSize,3) & ">3 (12 pt)</option>" & strLE & _
		"<option value=""4""" & chkSelect(strDefaultFontSize,4) & ">4 (14 pt)</option>" & strLE & _
		"<option value=""5""" & chkSelect(strDefaultFontSize,5) & ">5 (18 pt)</option>" & strLE & _
		"<option value=""6""" & chkSelect(strDefaultFontSize,6) & ">6 (24 pt)</option>" & strLE & _
		"<option value=""7""" & chkSelect(strDefaultFontSize,7) & ">7 (36 pt)</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Header Font Size</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strHeaderFontSize"">" & strLE & _
		"<option value=""""" & chkSelect(strHeaderFontSize,"") & ">None (blank)</option>" & strLE & _
		"<option value=""1""" & chkSelect(strHeaderFontSize,1) & ">1 (8 pt)</option>" & strLE & _
		"<option value=""2""" & chkSelect(strHeaderFontSize,2) & ">2 (10 pt)</option>" & strLE & _
		"<option value=""3""" & chkSelect(strHeaderFontSize,3) & ">3 (12 pt)</option>" & strLE & _
		"<option value=""4""" & chkSelect(strHeaderFontSize,4) & ">4 (14 pt)</option>" & strLE & _
		"<option value=""5""" & chkSelect(strHeaderFontSize,5) & ">5 (18 pt)</option>" & strLE & _
		"<option value=""6""" & chkSelect(strHeaderFontSize,6) & ">6 (24 pt)</option>" & strLE & _
		"<option value=""7""" & chkSelect(strHeaderFontSize,7) & ">7 (36 pt)</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Footer Font Size</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strFooterFontSize"">" & strLE & _
		"<option value=""""" & chkSelect(strFooterFontSize,"") & ">None (blank)</option>" & strLE & _
		"<option value=""1""" & chkSelect(strFooterFontSize,1) & ">1 (8 pt)</option>" & strLE & _
		"<option value=""2""" & chkSelect(strFooterFontSize,2) & ">2 (10 pt)</option>" & strLE & _
		"<option value=""3""" & chkSelect(strFooterFontSize,3) & ">3 (12 pt)</option>" & strLE & _
		"<option value=""4""" & chkSelect(strFooterFontSize,4) & ">4 (14 pt)</option>" & strLE & _
		"<option value=""5""" & chkSelect(strFooterFontSize,5) & ">5 (18 pt)</option>" & strLE & _
		"<option value=""6""" & chkSelect(strFooterFontSize,6) & ">6 (24 pt)</option>" & strLE & _
		"<option value=""7""" & chkSelect(strFooterFontSize,7) & ">7 (36 pt)</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontsize')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Base Background Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPageBGColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPageBGColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Default Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strDefaultFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strDefaultFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strLinkTextDecoration"">" & strLE & _
		"<option" & chkSelect(strLinkTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strLinkTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strLinkTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strLinkTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strLinkTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Visited Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strVisitedLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strVisitedLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Visited Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strVisitedTextDecoration"">" & strLE & _
		"<option" & chkSelect(strVisitedTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strVisitedTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strVisitedTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strVisitedTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strVisitedTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Active Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strActiveLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strActiveLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Active Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strActiveTextDecoration"">" & strLE & _
		"<option" & chkSelect(strActiveTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strActiveTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strActiveTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strActiveTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strActiveTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Hover Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strHoverFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHoverFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Hover Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strHoverTextDecoration"">" & strLE & _
		"<option" & chkSelect(strHoverTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strHoverTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strHoverTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strHoverTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strHoverTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Header Background Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strHeadCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHeadCellColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Header Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strHeadFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHeadFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Category Background Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strCategoryCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strCategoryCellColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Category Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strCategoryFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strCategoryFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>First Cell Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumFirstCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumFirstCellColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>First Alternating Cell Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumCellColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Second Alternating Cell Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strAltForumCellColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strAltForumCellColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strForumLinkTextDecoration"">" & strLE & _
		"<option" & chkSelect(strForumLinkTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strForumLinkTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strForumLinkTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strForumLinkTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strForumLinkTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Visited Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumVisitedLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumVisitedLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Visited Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strForumVisitedTextDecoration"">" & strLE & _
		"<option" & chkSelect(strForumVisitedTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strForumVisitedTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strForumVisitedTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strForumVisitedTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strForumVisitedTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Active Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumActiveLinkColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumActiveLinkColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Active Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strForumActiveTextDecoration"">" & strLE & _
		"<option" & chkSelect(strForumActiveTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strForumActiveTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strForumActiveTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strForumActiveTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strForumActiveTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Hover Link Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strForumHoverFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strForumHoverFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>In Forum Hover Link Decoration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strForumHoverTextDecoration"">" & strLE & _
		"<option" & chkSelect(strForumHoverTextDecoration,"none") & ">none</option>" & strLE & _
		"<option" & chkSelect(strForumHoverTextDecoration,"blink") & ">blink</option>" & strLE & _
		"<option" & chkSelect(strForumHoverTextDecoration,"line-through") & ">line-through</option>" & strLE & _
		"<option" & chkSelect(strForumHoverTextDecoration,"overline") & ">overline</option>" & strLE & _
		"<option" & chkSelect(strForumHoverTextDecoration,"underline") & ">underline</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#fontdecorations')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Table Border Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strTableBorderColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strTableBorderColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Pop-Up Table Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPopUpTableColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPopUpTableColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Pop-Up Table Border Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPopUpBorderColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strPopUpBorderColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>New Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strNewFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strNewFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>HighLight Font Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strHiLiteFontColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strHiLiteFontColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Search HighLight Color</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strSearchHiLiteColor"" size=""10"" maxLength=""20"" value=""" & chkExist(strSearchHiLiteColor) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#colors')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Page Background Image URL</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strPageBGImageURL"" size=""25"" maxLength=""100"" value=""" & chkExist(strPageBGImageURL) & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#pagebgimage')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Table Size Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Topic Left Column Width</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strTopicWidthLeft"" size=""5"" maxLength=""4"" value=""" & chkExistElse(strTopicWidthLeft,"100") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#columnwidth')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Topic NOWRAP Left</b>&nbsp;</td>" & strLE & _
		"<td><span class=""dff dfs"">" & strLE & _
		"On: <input type=""radio"" class=""radio"" name=""strTopicNoWrapLeft"" value=""1""" & chkRadio(strTopicNoWrapLeft,0,false) & ">" & strLE & _
		"Off: <input type=""radio"" class=""radio"" name=""strTopicNoWrapLeft"" value=""0""" & chkRadio(strTopicNoWrapLeft,0,true) & ">" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#nowrap')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Topic Right Column Width</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strTopicWidthRight"" size=""5"" maxLength=""4"" value=""" & chkExistElse(strTopicWidthRight,"100%") & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#columnwidth')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Topic NOWRAP Right</b>&nbsp;</td>" & strLE & _
		"<td><span class=""dff dfs"">" & strLE & _
		"On: <input type=""radio"" class=""radio"" name=""strTopicNoWrapRight"" value=""1""" & chkRadio(strTopicNoWrapRight,0,false) & ">" & strLE & _
		"Off: <input type=""radio"" class=""radio"" name=""strTopicNoWrapRight"" value=""0""" & chkRadio(strTopicNoWrapRight,0,true) & ">" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=colors#nowrap')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></span></td>" & strLE & _
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
