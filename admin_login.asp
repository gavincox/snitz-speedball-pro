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
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"Admin Login","class=""vam""") & "&nbsp;Admin&nbsp;Login" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE
fName         = strDBNTFUserName
fPassword     = ChkString(Request.Form("Password"), "SQLString")
RequestMethod = Request.ServerVariables("Request_method")
strTarget     = trim(chkString(request("target"),"SQLString"))
if RequestMethod = "POST" Then
	strEncodedPassword = sha256("" & fPassword)
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & trim(fName) & "' AND "
	strSql = strSql & " M_PASSWORD = '" & trim(strEncodedPassword) & "' AND "
	strSql = strSql & " M_LEVEL = 3 AND M_STATUS = 1"
	Set dbRs = my_Conn.Execute(strSql)
	If not(dbRS.EOF) and ChkQuoteOk(fName) and ChkQuoteOk(strEncodedPassword) Then
		Response.Write "<p class=""c""><span class=""dff hfs"">Login was successful!</span></p>" & strLE
		Session(strCookieURL & "Approval") = "15916941253"
		Response.Write "<p class=""c""><span class=""dff dfs""><a href="""
		if strTarget = "" then Response.Write "admin_home.asp" else Response.Write strTarget
		Response.Write """>Click here to Continue</a></span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL="
		if strTarget = "" then Response.Write "admin_home.asp" else Response.Write strTarget
		Response.Write """>" & strLE
		Call WriteFooter
		Response.End
	else
		Response.Write "<center>" & strLE & _
			"<p class=""c""><span class=""dff hfs hlfc"">There has been a problem!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff hfs hlfc"">You are not allowed access</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">If you think you have reached this message in error, please try again</span></p>" & strLE & _
			"</center>" & strLE
	end if
end if
Response.Write "<form action=""admin_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
	"<input type=""hidden"" value=""" & strTarget & """ name=""target"">" & strLE & _
	"<table class=""admin"">" & strLE & _
	"<tr>" & strLE & _
	"<th colspan=""2""><b>Admin Login</b></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw r""><b><span class=""dff dfs"">&nbsp;Username&nbsp;</span></b></td>" & strLE & _
	"<td><input type=""text"" name=""Name"" style=""width:150px;""></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw r""><b><span class=""dff dfs"">Password&nbsp;</span></b></td>" & strLE & _
	"<td><input type=""Password"" name=""Password"" style=""width:150px;""></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td colspan=""2"" class=""c""><input type=""submit"" value=""Login"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</form>" & strLE
Call WriteFooter
%>
