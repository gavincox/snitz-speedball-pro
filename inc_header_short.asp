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
<!--#INCLUDE FILE="inc_func_common.asp" -->
<%
strArchiveTablePrefix = strTablePrefix & "A_"
scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
strReferer = chkString(request.servervariables("HTTP_REFERER"),"refer")

set my_Conn= Server.CreateObject("ADODB.Connection")
my_Conn.Open strConnString

strDBNTUserName = Request.Cookies(strUniqueID & "User")("Name")
strDBNTFUserName = trim(chkString(Request.Form("Name"),"SQLString"))
if strDBNTFUserName = "" then strDBNTFUserName = trim(chkString(Request.Form("User"),"SQLString"))
if strAuthType = "nt" then
	strDBNTUserName = Session(strCookieURL & "userID")
	strDBNTFUserName = Session(strCookieURL & "userID")
end if

chkCookie = 1
mLev = cLng(chkUser(strDBNTUserName, Request.Cookies(strUniqueID & "User")("Pword"),-1))
chkCookie = 0

Response.Write "<!doctype html>" & strLE & _
	"<html lang=""en"">" & strLE & strLE & _
	"<head>" & strLE & _
	"<meta charset=""utf-8"">" & strLE & _
	"<title>" & chkString(strForumTitle,"pagetitle") & "</title>" & strLE
'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write  _
	"<meta name=""description"" content="""">" & strLE & _
	"<meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & strLE & _
	"<link href=""css/snitz.css"" rel=""stylesheet"" media=""all"">" & strLE & _
	"</head>" & strLE & strLE & _
	"<body onLoad=""window.focus();"">" & strLE
%>
