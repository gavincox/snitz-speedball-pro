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
<%
if MemberID > 0 then Response.Redirect("default.asp")
Response.Write "<table class=""tc"" width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw l"" width=""33%""><span class=""dff dfs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member&nbsp;Login<br></span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE

fName         = strDBNTFUserName
fPassword     = ChkString(Request.Form("Password"), "SQLString")
RequestMethod = Request.ServerVariables("Request_Method")
strTarget     = trim(chkString(request("target"),"SQLString"))

if RequestMethod = "POST" Then
	strEncodedPassword = sha256("" & fPassword)
	select case chkUser(fName, strEncodedPassword,-1)
		case 1, 2, 3, 4
			Call DoCookies(Request.Form("SavePassword"))
			strLoginStatus = 1
		case else : strLoginStatus = 0
	end select
	if strLoginStatus = 1 then
		Response.Write "<p class=""c""><span class=""dff hfs"">Login was successful!</span></p>" & strLE
		Response.Write "<p class=""c""><span class=""dff dfs""><a href="""
		if strTarget = "" then Response.Write "default.asp" else Response.Write strTarget
		Response.Write """>Click here to Continue</a></span></p>" & strLE
		Response.Write "<meta http-equiv=""Refresh"" content=""2; URL="
		if strTarget = "" then Response.Write "default.asp" else Response.Write strTarget
		Response.Write """>" & strLE & _
				"<br>" & strLE
		Call WriteFooter
		Response.End
	end if
end if
Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<form action=""login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
	"<input type=""hidden"" value=""" & strTarget & """ name=""target"">" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc tc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""hcc c""><b><span class=""dff dfs hfc"">Member Login</span></b></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc l""><b><span class=""dff dfs cfc"">Member Login</span></b></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""fcc"">" & strLE & _
	"<table class=""tc"" cellpadding=""6"" cellspacing=""0"" width=""90%"">" & strLE & _
	"<tr class=""vat"">" & strLE & _
	"<td width=""49%""><span class=""dff dfs ffc"">" & strLE
if RequestMethod = "POST" and strLoginStatus = 0 then Response.Write("<span class=""dff dfs hlfc"">Your username and/or password was incorrect.</span><br>" & strLE) else Response.Write("<br>" & strLE)
Response.Write "<b>Login:</b></span>" & strLE & _
	"<table cellpadding=""2"" cellspacing=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	" Username:<br>" & strLE & _
	"<input type=""text"" name=""Name"" size=""20"" maxLength=""25"" tabindex=""1"" value="""" style=""width:150px;""></td>" & strLE & _
"<td class=""vab"" rowspan=""2""><span class=""dff dfs ffc"">" & strLE
if strGfxButtons = "1" then
	Response.Write "<input src=""" & strImageUrl & "button_login.gif"" type=""image"" value=""Login"" id=""submit1"" name=""submit1"" tabindex=""3""></span></td>" & strLE
else
	Response.Write "<input class=""button"" type=""submit"" value=""Login"" id=""submit1"" name=""submit1"" tabindex=""3""></span></td>" & strLE
end if
Response.Write "</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	" Password:<br>" & strLE & _
	"<input type=""password"" name=""Password"" size=""20"" tabindex=""2"" maxLength=""25"" value="""" style=""width:150px;""></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<input type=""checkbox"" name=""SavePassWord"" tabindex=""4"" value=""true"" checked> Save Password</span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"<script type=""text/javascript"">document.Form1.Name.focus();</script>" & strLE & _
	"<td class=""nw"" width=""2%""></td>" & strLE & _
	"<td width=""49%""><span class=""dff dfs ffc""><br><b>Login Questions:</b><br>" & strLE & _
	"<span style=""font-size: 6px;""><br></span>" & strLE & _
	"<acronym title=""Do I have to register?""><span class=""smt""><a href=""faq.asp#register"">Do I have to register?</a></span></acronym><br>" & strLE
if strEmail = "1" then Response.Write("<acronym title=""Choose a new password if you have forgotten your current one.""><span class=""smt""><a href=""password.asp"">Forgot your Password?</a></span></acronym><br><br>" & strLE) else Response.Write("<br>" & strLE)
Response.Write " Not a member?<br>"
if strProhibitNewMembers = "1" then
	Response.Write "<span class=""ffs hlfc"">The Administrator has turned off registration for this forum<br>Only registered members are able to log in</span></span></td>" & strLE
else
	Response.Write "<acronym title=""Click here to register.""><span class=""smt""><a href=""register.asp"">Register Here!</a></span></acronymn></span></td>" & strLE
end if
Response.Write "</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</form>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE
Call WriteFooter
%>
