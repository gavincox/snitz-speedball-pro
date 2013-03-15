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
<!--#INCLUDE FILE="cb/admin_emaillist_cb.asp" -->
<%
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;User&nbsp;E-mail&nbsp;List<br><br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages c"">" & strLE
Call DropDownPaging(1)
Response.Write "</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE & _
	"<table class=""admin"">" & strLE & _
	"<caption><b>NOTE:</b> The following table will show you a list of all users of this forum, and their e-mail addresses.</caption>" & strLE & _
	"<tr>" & strLE & _
	"<th><b>User Name</b></th>" & strLE & _
	"<th><b>E-mail Address</b></th>" & strLE & _
	"<th class=""c""><b>Posts</b></th>" & strLE & _
	"</tr>" & strLE
if iMemberCount = "" then '## No Members Found in DB
	Response.Write "<tr>" & strLE & _
		"<td colspan=""3""><b>No Members Found</b></td>" & strLE & _
		"</tr>" & strLE
else
	mM_NAME  = 0
	mM_EMAIL = 1
	mM_POSTS = 2
	rec      = 1
	intI     = 0
	for iMember = 0 to iMemberCount
		if (rec = strPageSize + 1) then exit for
		Members_MemberName  = arrMemberData(mM_NAME, iMember)
		Members_MemberEMail = arrMemberData(mM_EMAIL, iMember)
		Members_MemberPosts = arrMemberData(mM_POSTS, iMember)
		Response.Write "<tr>" & strLE & _
			"<td>" & Members_MemberName & "</td>" & strLE & _
			"<td class=""smt""><a href=""mailto:" & Members_MemberEMail & """>" & Members_MemberEMail & "</a></td>" & strLE & _
			"<td class=""c"">" & Members_MemberPosts & "</td>" & strLE & _
			"</tr>" & strLE
		rec  = rec + 1
		intI = intI + 1
		if intI = 2 then intI = 0
	next
end if
Response.Write "</table>" & strLE & _
	"<div id=""post-content"">" & strLE & _
"<div class=""maxpages c"">" & strLE
Call DropDownPaging(1)
Response.Write "</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /post-content -->" & strLE
Call WriteFooter
Response.End
%>
