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
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Ranking&nbsp;Configuration" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strRankAdmin") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Administrator Name</li>"
	'if Request.Form("strRankGlobalMod") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Global Moderator Name</li>"
 	if Request.Form("strRankMod") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Moderator Name</li>"
	if Request.Form("strRankLevel0") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Starting Member Name</li>"
	if Request.Form("strRankLevel1") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 1 Name</li>"
	if Request.Form("strRankLevel2") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 2 Name</li>"
	if Request.Form("strRankLevel3") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 3 Name</li>"
	if Request.Form("strRankLevel4") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 4 Name</li>"
	if Request.Form("strRankLevel5") = "" then Err_Msg = Err_Msg & "<li>You Must Enter a Value for Member Level 5 Name</li>"
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel2")) then Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 2</li>"
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel3")) then Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 3</li>"
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel3")) then Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 3</li>"
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel4")) then Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 4</li>"
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel4")) then Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 4</li>"
	if cLng(Request.Form("intRankLevel3")) > cLng(Request.Form("intRankLevel4")) then Err_Msg = Err_Msg & "<li>Rank Level 3 can not be higher than 4</li>"
	if cLng(Request.Form("intRankLevel1")) > cLng(Request.Form("intRankLevel5")) then Err_Msg = Err_Msg & "<li>Rank Level 1 can not be higher than 5</li>"
	if cLng(Request.Form("intRankLevel2")) > cLng(Request.Form("intRankLevel5")) then Err_Msg = Err_Msg & "<li>Rank Level 2 can not be higher than 5</li>"
	if cLng(Request.Form("intRankLevel3")) > cLng(Request.Form("intRankLevel5")) then Err_Msg = Err_Msg & "<li>Rank Level 3 can not be higher than 5</li>"
	if cLng(Request.Form("intRankLevel4")) > cLng(Request.Form("intRankLevel5")) then Err_Msg = Err_Msg & "<li>Rank Level 4 can not be higher than 5</li>"
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
		"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
	end if
else
	arrStarColors = ("gold|silver|bronze|orange|red|purple|blue|cyan|green")
	arrIconStarColors = array(strIconStarGold,strIconStarSilver,strIconStarBronze,strIconStarOrange,strIconStarRed,strIconStarPurple,strIconStarBlue,strIconStarCyan,strIconStarGreen)
	strStarColor = split(arrStarColors, "|")
	Response.Write "<form action=""admin_config_ranks.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>Ranking Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Show Ranking</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strShowRank"">" & strLE & _
		"<option value=""0""" & chkSelect(strShowRank,0) & ">None</option>" & strLE & _
		"<option value=""1""" & chkSelect(strShowRank,1) & ">Rank Only</option>" & strLE & _
		"<option value=""2""" & chkSelect(strShowRank,2) & ">Stars Only</option>" & strLE & _
		"<option value=""3""" & chkSelect(strShowRank,3) & ">Rank and Stars</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#ShowRank')"">" & getCurrentIcon(strIconSmileQuestion,"ShowRank","") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Administrator Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankAdmin"" size=""30"" value=""" & chkExistElse(chkString(strRankAdmin,"edit"),"Administrator") & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Administrator)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColorAdmin"" value=""" & strStarColor(c) & """" & chkRadio(strRankColorAdmin,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Moderator Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankMod"" size=""30"" value=""" & chkExistElse(chkString(strRankMod,"edit"),"Moderator") & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Moderator)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColorMod"" value=""" & strStarColor(c) & """" & chkRadio(strRankColorMod,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Starting Member Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel0"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel0,"edit"),"Starting Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel0"" size=""5"" value=""0"" readonly>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has less than Member Level 1 but more than Starting Member Level posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Member Level 1 Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel1"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel1,"edit"),"New Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel1"" size=""5"" value=""" & chkExistElse(intRankLevel1,50) & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 1 and Member Level 2 posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColor1"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor1,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Member Level 2 Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel2"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel2,"edit"),"Junior Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel2"" size=""5"" value=""" & chkExistElse(intRankLevel2,100) & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 2 and Member Level 3 posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColor2"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor2,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Member Level 3 Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel3"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel3,"edit"),"Average Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel3"" size=""5"" value=""" & chkExistElse(intRankLevel3,500) & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 3 and Member Level 4 posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColor3"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor3,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Member Level 4 Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel4"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel4,"edit"),"Senior Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel4"" size=""5"" value=""" & chkExistElse(intRankLevel4,1000) & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has between Member Level 4 and Member Level 5 posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColor4"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor4,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Member Level 5 Name</b>&nbsp;</td>" & strLE & _
		"<td><input type=""text"" name=""strRankLevel5"" size=""30"" value=""" & chkExistElse(chkString(strRankLevel5,"edit"),"Advanced Member") & """>" & strLE & _
		"<b>Number</b>&nbsp;<input type=""text"" name=""intRankLevel5"" size=""5"" value=""" & chkExistElse(intRankLevel5,2000) & """>" & strLE & _
		getCurrentIcon(strIconSmileQuestion,"(Member who has more than Member Level 5 posts)","class=""vam""") & "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Star Color</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE
	for c = 0 to ubound(strStarColor)
		Response.Write "<input type=""radio"" class=""radio"" name=""strRankColor5"" value=""" & strStarColor(c) & """" & chkRadio(strRankColor5,strStarColor(c),true) & ">&nbsp;" & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & strLE
	next
	Response.Write "<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=ranks#RankColor')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a>&nbsp;</td>" & strLE & _
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
