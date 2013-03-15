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
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
		getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Forum&nbsp;Variables&nbsp;Information<br><br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE & _
	"<table class=""admin"" style=""width:300px"">" & strLE & _
		"<caption><b>NOTE:</b> The following table will show you values of the different variables used by the Forum</caption>" & strLE & _
		"<tr>" & strLE & _
		"<th><b>Variable&nbsp;Name</b></th>" & strLE & _
		"<th><b>Value</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr><th colspan=""2""><b>General&nbsp;information</b></td></tr>" & strLE & _
		"<tr><td><b>strCookieUrl</b></td><td>" & ChkString(StrCookieUrl, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>strUniqueID</b></td><td>" & ChkString(StrUniqueID, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>strAuthType</b></td><td>" & ChkString(strAuthType, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>strDBNTSQLName</b></td><td>" & ChkString(strDBNTSQLName, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>strDBNTUserName</b></td><td>" & ChkString(strDBNTUserName, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>strDBType</b></td><td>" & ChkString(strDBType, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><td><b>intCookieDuration</b></td><td>" & ChkString(intCookieDuration, "admindisplay") & "</td></tr>" & strLE & _
		"<tr><th colspan=""2""><b>Cookies</b></th>" & strLE & _
		"</tr>" & strLE
for each key in Request.Cookies
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		if Request.Cookies(key).HasKeys then
			for each subkey in Request.Cookies(key)
				Response.Write "<tr><td class=""vat""><b>" & chkString(key, "admindisplay") & " (" & chkString(subkey, "admindisplay") & ")</b></td>" & strLE & _
					"<td><span face=""courier"">"
				if Request.Cookies(key)(subkey) = "" then
					Response.Write "&nbsp;"
				else
					Response.Write ChkString(CStr(Request.Cookies(key)(subkey)), "admindisplay")
				end if
				Response.Write "</td></tr>" & strLE
			next
		else
			Response.Write "<tr><td class=""vat""><b>" & chkString(key, "admindisplay") & "</b></td>" & strLE & _
				"<td><span face=""courier"">"
			if Request.Cookies(key) = "" then
				Response.Write "&nbsp;"
			else
				Response.Write ChkString(CStr(Request.Cookies(key)), "admindisplay")
			end if
			Response.Write "</td></tr>" & strLE
		end if
	end if
next
Response.Write "<tr>" & strLE & _
		"<th colspan=""2""><b>Session&nbsp;variables</b></th></tr>" & strLE
for each key in Session.Contents
	if not IsArray(Session.Contents(key)) then
		if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
			Response.Write "<tr>" & strLE & _
				"<td class=""vat""><b>" & ChkString(key, "admindisplay") & "</b></td>" & strLE & _
				"<td><span face=""courier"">"
			if Session.Contents(key) = "" then
				Response.Write "&nbsp;"
			else
				Response.Write chkString(CStr(Session.Contents(key)), "admindisplay")
			end if
			Response.Write "</td></tr>" & strLE
		end if
	end if
next
Response.Write "<tr>" & strLE & _
		"<th colspan=""2""><b>Application&nbsp;variables</b></th></tr>" & strLE
for each key in Application.Contents
	if left(lcase(key), len(strCookieUrl)) = lcase(strCookieUrl) or left(lcase(key), len(strUniqueID)) = lcase(strUniqueID) then
		Response.Write "<tr>" & strLE & _
			"<td class=""putc vat""><span class=""dff dfs""><b>" & chkString(key, "admindisplay") & "</b></span></td>" & strLE & _
			"<td class=""putc""><span face=""courier"" class=""dfs"">"
		if Application.Contents(key) = "" then
			Response.Write "&nbsp;"
		else
			Response.Write chkString(CStr(Application.Contents(key)), "admindisplay")
		end if
		Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE
	end if
next
Response.Write "</table>" & strLE
Call WriteFooter
Response.End
%>
