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
<!--#INCLUDE FILE="cb/admin_accounts_pending_cb.asp" -->
<%
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs w50"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Members&nbsp;Pending" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages w50"">" & strLE
if maxpages > 1 then Call DropDownPaging(1)
Response.Write "</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content-->" & strLE
if iMemberCount <> "" then
	if strRestrictReg = "1" then scolspan = " colspan=""2"""
	Response.Write "<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th><b>Administrator Options</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE
	if strRestrictReg = "1" then
		Response.Write "<td class=""fcc""><ul>" & strLE & _
			"<li><a href=""javascript:appr_all()"">Approve All Pending Members</a></li>" & strLE & _
			"<li><a href=""javascript:appr_selected()"">Approve Selected Pending Members</a></li></ul></td>" & strLE
	end if
	Response.Write "<td class=""fcc""><ul>" & strLE & _
		"<li><a href=""javascript:del_all()"">Delete All Pending Members</a></li>" & strLE & _
		"<li><a href=""javascript:del_selected()"">Delete Selected Pending Members</a></li></ul></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table><br>" & strLE
end if
if iMemberCount <> "" then
	Response.Write "<form name=""delMembers"" action=""admin_accounts_pending.asp"">" & strLE & _
		"<input type=""hidden"" name=""action"" value=""none"">" & strLE & _
		"<input type=""hidden"" name=""whichpage"" value=""" & mypage & """>" & strLE
end if
Response.Write "<table id=""content"">" & strLE & _
	"<tbody>" & strLE & _
	"<tr>" & strLE & _
	"<caption class=""l""><b>NOTE:</b> The following table will show you a list of registered users that are waiting to be authenticated</caption>" & strLE & _
	"<tr>" & strLE & _
	"<th><b>User Name</b></th>" & strLE & _
	"<th><b>E-mail Address</b></th>" & strLE & _
	"<th><b>IP Address</b></th>" & strLE & _
	"<th><b>Registered</b></th>" & strLE & _
	"<th><b>Days Since</b></th>" & strLE & _
	"<th><b>Action</b></th>" & strLE
if strRestrictReg = "1" then Response.Write "<th><b>Approved?</b></th>" & strLE
Response.Write "<th><b>"
if iMemberCount <> "" then
	Response.Write "<input type=""checkbox"" name=""toggleAll"" value="""" onClick=""ToggleAll(this);"">"
else
	Response.Write "&nbsp;"
end if
Response.Write "</b></th>" & strLE & _
	"</tr>" & strLE
if iMemberCount = "" then  '## No members found in DB
	if strRestrictReg = "1" then intcolspan = 8 else intcolspan = 7
	Response.Write "<tr>" & strLE & _
		"<td class=""fcc"" colspan=""" & intcolspan & """><b>No Members Found</b></td>" & strLE & _
		"</tr>" & strLE
else
	mM_NAME    = 0
	mM_EMAIL   = 1
	mMEMBER_ID = 2
	mM_DATE    = 3
	mM_IP      = 4
	mM_KEY     = 5
	mM_APPROVE = 6
	rec  = 1
	intI = 0
	for iMember = 0 to iMemberCount
		if (rec = strPageSize + 1) then exit for
		MP_MemberName    = arrMemberData(mM_NAME, iMember)
		MP_MemberEMail   = arrMemberData(mM_EMAIL, iMember)
		MP_MemberID      = arrMemberData(mMEMBER_ID, iMember)
		MP_MemberDate    = arrMemberData(mM_DATE, iMember)
		MP_MemberIP      = arrMemberData(mM_IP, iMember)
		MP_MemberKey     = arrMemberData(mM_KEY, iMember)
		MP_MemberApprove = arrMemberData(mM_APPROVE, iMember)
		if intI = 1 then
			CColor = strAltForumCellColor
		else
			CColor = strForumCellColor
		end if
		if MP_MemberApprove = 1 then
			Approved = "Yes"
		else
			Approved = "No"
		end if
		days = DateDiff("d",  StrToDate(MP_MemberDate),  strForumTimeAdjust)
		if days >= 15 then
			days2 = "<b>" & days & "</b>"
		else
			days2 = days
		end if
		Response.Write "<tr>" & strLE & _
			"<td class=""c""><a href=""pop_profile_pending.asp?mode=display&id="& MP_MemberID & """>"& chkString(MP_MemberName, "display") & "</a></td>" & strLE & _
			"<td class=""ffs c"">" & MP_MemberEMail & "</td>" & strLE & _
			"<td class=""c""><a href=""" & strIPLookup & ChkString(MP_MemberIP, "display") & """ target=""_blank"">" & MP_MemberIP & "</a></td>" & strLE & _
			"<td class=""ffs c"">" & ChkDate(MP_MemberDate,"<br>",true) & "</td>" & strLE & _
			"<td class=""c"">"
		if days >= 7 then Response.Write(" hlfc") else Response.Write " ffc"
		Response.Write """>" & days2 & "</td>" & strLE & _
			"<td class=""ffs smt c""><a href=""register.asp?actkey=" & MP_MemberKey & """>Activate Account</a></td>" & strLE
		if strRestrictReg = "1" then Response.Write "<td class=""c"">" & Approved & "</td>" & strLE
		Response.Write "<td class=""ffs c""><input type=""checkbox"" name=""id"" value=""" & MP_MemberID & """ onclick=""Toggle(this)""></td>" & strLE & _
			"</tr>" & strLE
		rec  = rec + 1
		intI = intI + 1
		if intI = 2 then intI = 0
	next
	Response.Write "</form>"
end if
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
if maxpages > 1 then
	Response.Write "<table class=""tl"">" & strLE & _
		"<tr>" & strLE
	Call DropDownPaging(2)
	Response.Write "</tr>" & strLE & _
		"</table>" & strLE
else
	Response.Write "<br>" & strLE
end if
Call WriteFooter
Response.End
%>