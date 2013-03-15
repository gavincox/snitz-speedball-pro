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
if strAuthType <> "nt" then
	Response.Redirect "admin_home.asp"
end if
Response.Write "<table width=""100%"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""nw l"" width=""33%""><span class=""dff dfs"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
		getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Feature&nbsp;NT&nbsp;Configuration<br></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE
if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strIMGInPosts") = "1" and Request.Form("strAllowForumCode") = "0" then
		Err_Msg = Err_Msg & "<li>Forum Code Must be Enabled in order to Enable Images</li>"
	end if
	if (Request.Form("strHotTopic") = "1" and strHotTopic = "1") or (Request.Form("strHotTopic") = "1" and strHotTopic = "0") then
		if Request.Form("intHotTopicNum") = "" then
			Err_Msg = Err_Msg & "<li>You Must Enter a Hot Topic Number</li>"
		end if
		if left(Request.Form("intHotTopicNum"), 1) = "-" then
			Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number</li>"
		end if
		if left(Request.Form("intHotTopicNum"), 1) = "+" then
			Err_Msg = Err_Msg & "<li>You Must Enter a positive Hot Topic Number without the <b>+</b></li>"
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
				"<p class=""c""><span class=""dff dfs"">Congratulations!</span></p>" & strLE & _
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
	Response.Write "<form action=""admin_config_NT_features.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
			"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
			"<table class=""tc"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""pubc"">" & strLE & _
			"<table cellspacing=""1"" cellpadding=""1"">" & strLE & _
			"<tr class=""vat"">" & strLE & _
			"<td class=""hcc"" colspan=""2""><span class=""dff dfs hfc""><b>Feature NT Configuration</b></span></td>" & strLE & _
			"</tr>" & strLE
	if strAuthType = "nt" then
		Response.Write "<tr class=""vat"">" & strLE & _
				"<td class=""putc r""><span class=""dff dfs""><b>Use NT Groups:</b>&nbsp;</span></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs"">" & strLE & _
				"                On: <input type=""radio"" class=""radio"" name=""strNTGroups"" value=""1""" & chkRadio(strNTGroups,0,false) & ">&nbsp;" & strLE & _
				"                Off: <input type=""radio"" class=""radio"" name=""strNTGroups"" value=""0""" & chkRadio(strNTGroups,0,true) & ">" & strLE & _
				"<a href=""JavaScript:openWindow3('pop_config_help.asp')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strAuthType = "nt" then
		Response.Write "<tr class=""vat"">" & strLE & _
				"<td class=""putc r""><span class=""dff dfs""><b>Use NT AutoLogon:</b>&nbsp;</span></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs"">" & strLE & _
				"                On: <input type=""radio"" class=""radio"" name=""strAutoLogon"" value=""1""" & chkRadio(strAutoLogon,0,false) & ">&nbsp;" & strLE & _
				"                Off: <input type=""radio"" class=""radio"" name=""strAutoLogon"" value=""0""" & chkRadio(strAutoLogon,0,true) & ">" & strLE & _
				"<a href=""JavaScript:openWindow3('pop_config_help.asp')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	Response.Write "<tr class=""vat"">" & strLE & _
			"<td class=""putc c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" id=""submit1"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE
end if
WriteFooter
Response.End
%>
