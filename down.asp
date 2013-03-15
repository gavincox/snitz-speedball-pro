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
<!--#INCLUDE FILE="config.asp" -->
<!--#INCLUDE FILE="inc_func_common.asp" -->
<%
Dim status, info1, info2, fStatus
mlev     = request("mlev")
status   = Application(strCookieURL & "down")
fStatus  = request.form("status")
DMessage = request.Form("DownMessage")

if DMessage = "" then DMessage = Application(strCookieURL & "DownMessage")
if status   = "" then status = false

if (not isEmpty(fStatus)) and (Session(strCookieURL & "Approval") = "15916941253") then
	if status then
		Application.lock
		Application(strCookieURL & "down") = false
		Application(strCookieURL & "DownMessage") = ""
		Application.unlock
		status = false
	else
		Application.lock
		Application(strCookieURL & "down") = true
		Application(strCookieURL & "DownMessage") = DMessage
		Application.unlock
		status = true
	end if
end if

if status then
	info1 = "down"
	info2 = "Start"
else
	info1 = "running"
	info2 = "Stop"
end if

if Session(strCookieURL & "Approval") = "15916941253" Then
	strScriptName = request.servervariables("script_name")
	Response.Write "<html>" & strLE & _
		"<head>" & strLE & _
		"<title>" & GetNewTitle(strScriptName) & "</title>" & strLE
	'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write "<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
	'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & strLE & _
		"<link href=""css/normalize-legacy.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"<link href=""css/snitz.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"</head>" & strLE & _
		"<body>" & strLE & strLE & _
		"<header role=""banner"">" & strLE & _
		"<div id=""header"">" & strLE & _
		"<div class=""hdlogo""><a href=""default.asp"">" & getCurrentIcon(strTitleImage & "||",strForumTitle,"") & "</a></div>" & strLE & _
		"<!-- /hdlogo -->" & strLE & _
		"<div class=""hdcontrols"">" & strLE & _
		"</div>" & strLE & _
		"<!-- /hdcontrols -->" & strLE & _
		"</div>" & strLE & _
		"<!-- /header -->" & strLE & _
		"</header>" & strLE & strLE & _
		"<div id=""pre-content"">" & strLE & _
		"<div class=""breadcrumbs"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
		getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Forum&nbsp;Maintenance<br></span></td>" & strLE & _
		"</div>" & strLE & _
		"<!-- /breadcrumbs -->" & strLE & _
		"</div>" & strLE & _
		"<!-- /pre-content -->" & strLE & _
		"<br><br><br>" & strLE & _
		"<div class=""lmessage"" style=""clear:both;margin-top:50px"">" & strLE & _
		"<form action=""down.asp"" method=""post"">" & strLE & _
		"<p><b>Welcome Administrator. The current status of the boards is <span class=""hlfc"">" & info1 & "</span>.</b></p>" & strLE & _
		"<p>would you like to:</p>" & strLE & _
		"<input type=""submit"" value=""" & info2 & " the board"" name=""Submit"">" & strLE & _
		"<input type=""hidden"" name=""status"" value=""" & status & """>" & strLE & _
		"<p>The message below will appear when the board is closed.</p>" & strLE & _
		"<textarea cols=""80"" rows=""12"" name=""DownMessage"" wrap=""soft"">" & Application(strCookieURL & "DownMessage") & "</textarea>" & strLE & _
		"</form>" & strLE & _
		"</div>" & strLE & _
		"<!-- /lmessage -->" & strLE & _
		"<br>" & strLE
else
	if mlev = 4 then
		Response.Redirect "admin_login.asp?target=down.asp?mlev=" & mLev
	elseif not Application(strCookieURL & "down") then
		response.redirect "default.asp"
	end if
	strScriptName = request.servervariables("script_name")
	Response.Write "<html>" & strLE & _
		"<head>" & strLE & _
		"<title>" & GetNewTitle(strScriptName) & "</title>" & strLE
	'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write "<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
	'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write "<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & strLE & _
		"<link href=""css/normalize-legacy.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"<link href=""css/snitz.css"" rel=""stylesheet"" media=""all"">" & strLE & _
		"</head>" & strLE & _
		"<body>" & strLE & strLE & _
		"<header role=""banner"">" & strLE & _
		"<div id=""header"">" & strLE & _
		"<div class=""hdlogo""><a href=""default.asp"">" & getCurrentIcon(strTitleImage & "||",strForumTitle,"") & "</a></div>" & strLE & _
		"<!-- /hdlogo -->" & strLE & _
		"<div class=""hdcontrols"">" & strLE & _
		"</div>" & strLE & _
		"<!-- /hdcontrols -->" & strLE & _
		"</div>" & strLE & _
		"<!-- /header -->" & strLE & _
		"</header>" & strLE & strLE & _
		"<div id=""pre-content""></div>" & strLE & _
		"<!-- /pre-content -->" & strLE & _
		"<br><br><br>" & strLE & _
		"<div class=""lmessage w95 tc"">" & strLE & _
		"<p>" & strForumTitle & " is currently closed.</p>" & strLE & _
		"<p>The Administrator has chosen to close<br>" & strLE & _
		"this forum with the following reason:<br><br>" & strLE & _
		"<b>" & Application(strCookieURL & "DownMessage") & "</b><br><br>" & strLE & _
		"<a href=""admin_login.asp?target=down.asp"">Administrator Login</a>" & strLE & _
		"</div>" & strLE
end if
Response.Write "</body>" & strLE & _
	"</html>" & strLE
%>
