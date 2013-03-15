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
on error resume next
strName  = my_Conn.Properties(0).name
strValue = my_Conn.Properties(0).value
on error goto 0
if Err.Number <> 0 then blnDisplay = False else blnDisplay = True
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Server&nbsp;Information<br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE & _
	"<table class=""admin"" style=""width:300px"">" & strLE & _
	"<caption><b>NOTE:</b> The following table will show you values of interest in setting up these forums. Most useful will be the line that shows the APPL_PHYSICAL_PATH. This can be used to properly write your DSN'less Connection String.</caption>" & strLE & _
	"<tr>" & strLE & _
	"<th><b>Variable&nbsp;Name</b></th>" & strLE & _
	"<th><b>Value</b></th>" & strLE & _
	"</tr>" & strLE
for each key in Request.ServerVariables
	Response.Write "<tr class=""vat"">" & strLE & _
		"<td><b>" & key & "</b></td>" & strLE & _
		"<td><span face=""courier"">"
	if Request.ServerVariables(key) = "" then
		Response.Write "&nbsp;"
	else
		Response.Write Request.Servervariables(key)
	end if
	Response.Write "</span></td>" & strLE & _
		"</tr>" & strLE
next
if blnDisplay = True then
	'## Code below added to show general ADO/Database Information
	Response.Write "<tr>" & strLE & _
		"<th colspan=""2""><b>Database Connection Properties</b></td>" & strLE & _
		"</tr>" & strLE
	for each item in my_Conn.Properties
		Response.Write "<tr>" & strLE & _
			"<td class=""vat""><b>" & item.name & "</b></td>" & strLE & _
			"<td><span face=""courier"">"
		if item.value = "" then Response.Write "&nbsp;" else Response.Write item.value
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE
	next
	'## Code above added to show general ADO/Database Information
end if
Response.Write "</table>" & strLE
Call WriteFooter
Response.End
%>
