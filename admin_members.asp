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
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
'## Forum_SQL - Get all members
strSql  = "SELECT MEMBER_ID, M_STATUS, M_NAME, M_LEVEL, M_EMAIL, M_TITLE, M_POSTS, M_LASTPOSTDATE, M_LASTHEREDATE, M_DATE "
strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
strSql3 = " WHERE M_LEVEL > 1 "
strSql4 = " ORDER BY M_LEVEL ASC, M_NAME ASC"
set rs  = Server.CreateObject("ADODB.Recordset")
rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
	if rs.EOF then
		iMemberCount = ""
	else
		arrMemberData = rs.GetRows(adGetRowsRest)
		iMemberCount = UBound(arrMemberData,2)
	end if
rs.Close
set rs = nothing
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Current&nbsp;Admins&nbsp;and&nbsp;Moderators" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE & _
	"<table class=""admin"">" & strLE & _
	"<tr>" & strLE & _
	"<th><b>Member</b></th>" & strLE & _
	"<th><b>Title</b></th>" & strLE & _
	"<th><b>Member Since</b></th>" & strLE & _
	"<th><b>Last Post</b></th>" & strLE & _
	"<th><b>Last Visit</b></th>" & strLE & _
	"<th><b>&nbsp;</b></th>" & strLE & _
	"</tr>" & strLE
if iMemberCount = "" then '## No Members Found in DB
	Response.Write "<tr>" & strLE & _
		"<td colspan=""6"" class=""fcc""><b>No Members Found</b></td>" & strLE & _
		"</tr>" & strLE
else
	mMEMBER_ID      = 0
	mM_STATUS       = 1
	mM_NAME         = 2
	mM_LEVEL        = 3
	mM_EMAIL        = 4
	mM_TITLE        = 5
	mM_POSTS        = 6
	mM_LASTPOSTDATE = 7
	mM_LASTHEREDATE = 8
	mM_DATE         = 9
	rec             = 1
	intI            = 0
	for iMember = 0 to iMemberCount
		Members_MemberID           = arrMemberData(mMEMBER_ID, iMember)
		Members_MemberStatus       = arrMemberData(mM_STATUS, iMember)
		Members_MemberName         = arrMemberData(mM_NAME, iMember)
		Members_MemberLevel        = arrMemberData(mM_LEVEL, iMember)
		Members_MemberEMail        = arrMemberData(mM_EMAIL, iMember)
		Members_MemberTitle        = arrMemberData(mM_TITLE, iMember)
		Members_MemberPosts        = arrMemberData(mM_POSTS, iMember)
		Members_MemberLastPostDate = arrMemberData(mM_LASTPOSTDATE, iMember)
		Members_MemberLastHereDate = arrMemberData(mM_LASTHEREDATE, iMember)
		Members_MemberDate         = arrMemberData(mM_DATE, iMember)
		if Members_MemberLevel = 3 then
			if rec = 1 then
				Response.Write "<tr>" & strLE & _
					"<th colspan=""6""><b>Administrators</b></th>" & strLE & _
					"</tr>" & strLE
			end if
			if intI = 1 then
				CColor = ffacc
			else
				CColor = fsacc
			end if
			Response.Write "<tr>" & strLE & _
				"<td class=""smt c"">" & profileLink(ChkString(Members_MemberName,"display"),Members_MemberID) & "</td>" & strLE & _
				"<td class=""ffs c"">" & ChkString(getMember_Level(Members_MemberTitle, Members_MemberLevel, Members_MemberPosts),"display") & "</td>" & strLE
			Response.Write "<td class=""ffs nw c"">" & ChkDate(Members_MemberDate,"",false) & "</td>" & strLE
			if IsNull(Members_MemberLastHereDate) or Trim(Members_MemberLastPostDate) = "" then
				Response.Write "<td class=""ffs nw c"">-</td>" & strLE
			else
				Response.Write "<td class=""ffs nw c"">" & ChkDate(Members_MemberLastPostDate,"<br>",true) & "</td>" & strLE
			end if
			Response.Write "<td class=""nw c""><span class=""dff ffs ffc"">" & ChkDate(Members_MemberLastHereDate,"<br>",true) & "</span></td>" & strLE
			Response.Write "<td class=""c""><b>" & strLE
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				if Members_MemberStatus <> 0 then
					Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconLock,"Lock Member","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (Members_MemberID = intAdminMemberID and MemberID <> intAdminMemberID) OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID AND MemberID <> Members_MemberID) then
				Response.Write "-" & strLE
			else
				if strUseExtendedProfile then
					Response.Write "<a href=""pop_profile.asp?mode=Modify&amp;ID=" & Members_MemberID & """>" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow3('pop_profile.asp?mode=Modify&amp;ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Member","class=""vam""") & "</a>" & strLE
			end if
			Response.Write "</b></td>" & strLE
			Response.Write "</tr>" & strLE
			rec = rec + 1
			intI = intI + 1
			if intI = 2 then intI = 0
		end if
	next
	rec  = 1
	intI = 0
	for iMember = 0 to iMemberCount
		Members_MemberID           = arrMemberData(mMEMBER_ID, iMember)
		Members_MemberStatus       = arrMemberData(mM_STATUS, iMember)
		Members_MemberName         = arrMemberData(mM_NAME, iMember)
		Members_MemberLevel        = arrMemberData(mM_LEVEL, iMember)
		Members_MemberEMail        = arrMemberData(mM_EMAIL, iMember)
		Members_MemberTitle        = arrMemberData(mM_TITLE, iMember)
		Members_MemberPosts        = arrMemberData(mM_POSTS, iMember)
		Members_MemberLastPostDate = arrMemberData(mM_LASTPOSTDATE, iMember)
		Members_MemberLastHereDate = arrMemberData(mM_LASTHEREDATE, iMember)
		Members_MemberDate         = arrMemberData(mM_DATE, iMember)
		if Members_MemberLevel = 2 then
			if rec = 1 then
				Response.Write "<tr>" & strLE & _
					"<th colspan=""6""><b>Moderators</b></th>" & strLE & _
					"</tr>" & strLE
			end if
			if intI = 1 then
				CColor = strAltForumCellColor
			else
				CColor = strForumCellColor
			end if
			Response.Write "<tr>" & strLE & _
				"<td class=""smt c"">" & profileLink(ChkString(Members_MemberName,"display"),Members_MemberID) & "</td>" & strLE & _
				"<td class=""ffs c"">" & ChkString(getMember_Level(Members_MemberTitle, Members_MemberLevel, Members_MemberPosts),"display") & "</td>" & strLE & _
				"<td class=""ffs nw c"">" & ChkDate(Members_MemberDate,"",false) & "</td>" & strLE
			if IsNull(Members_MemberLastHereDate) or Trim(Members_MemberLastPostDate) = "" then
				Response.Write "<td class=""ffs nw c"">-</td>" & strLE
			else
				Response.Write "<td class=""ffs nw c"">t" & ChkDate(Members_MemberLastPostDate,"<br>",true) & "</td>" & strLE
			end if
			Response.Write "<td class=""ffs nw c"">" & ChkDate(Members_MemberLastHereDate,"<br>",true) & "</td>" & strLE
			Response.Write "<td class=""c""><b>" & strLE
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				if Members_MemberStatus <> 0 then
					Response.Write "<a href=""JavaScript:openWindow('pop_lock.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconLock,"Lock Member","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow('pop_open.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconUnlock,"Un-Lock Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if (Members_MemberID = intAdminMemberID and MemberID <> intAdminMemberID) OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID AND MemberID <> Members_MemberID) then
				Response.Write " -" & strLE
			else
				if strUseExtendedProfile then
					Response.Write "<a href=""pop_profile.asp?mode=Modify&amp;ID=" & Members_MemberID & """>" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				else
					Response.Write "<a href=""JavaScript:openWindow3('pop_profile.asp?mode=Modify&amp;amp;ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconPencil,"Edit Member","class=""vam""") & "</a>" & strLE
				end if
			end if
			if Members_MemberID = intAdminMemberID OR (Members_MemberLevel = 3 AND MemberID <> intAdminMemberID) then
				'## Do Nothing
			else
				Response.Write "<a href=""JavaScript:openWindow('pop_delete.asp?mode=Member&amp;MEMBER_ID=" & Members_MemberID & "')"">" & getCurrentIcon(strIconTrashcan,"Delete Member","class=""vam""") & "</a>" & strLE
			end if
			Response.Write "</b></td>" & strLE
			Response.Write "</tr>" & strLE
			rec = rec + 1
			intI = intI + 1
			if intI = 2 then intI = 0
		end if
	next
end if
Response.Write "</table>" & strLE
Call WriteFooter
Response.End
%>
