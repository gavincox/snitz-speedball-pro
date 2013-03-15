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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Request("CAT_ID") <> "" then
	if IsNumeric(Request("CAT_ID")) = True then Cat_ID = cLng(Request("CAT_ID")) else Cat_ID = 0
end if
if Request("FORUM_ID") <> "" then
	if IsNumeric(Request("FORUM_ID")) = True then Forum_ID = cLng(Request("FORUM_ID")) else Forum_ID = 0
end if
if Request("TOPIC_ID") <> "" then
	if IsNumeric(Request("TOPIC_ID")) = True then Topic_ID = cLng(Request("TOPIC_ID")) else Topic_ID = 0
end if
if Request("REPLY_ID") <> "" then
	if IsNumeric(Request("REPLY_ID")) = True then Reply_ID = cLng(Request("REPLY_ID")) else Reply_ID = 0
end if
if Request("MEMBER_ID") <> "" then
	if IsNumeric(Request("MEMBER_ID")) = True then Member_ID = cLng(Request("MEMBER_ID")) else Member_ID = 0
end if

if (Cat_ID + Forum_ID + Topic_ID + Reply_ID + Member_ID) < 1 then
	Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>The URL has been modified!</b></span></p>" & strLE & _
		"<p class=""c""><span class=""dff dfs hlfc""><b>Possible Hacking Attempt!</b></span></p>" & strLE
	Call WriteFooterShort
	Response.End
end if

Mode_Type = ChkString(Request("mode"), "SQLString")
strPassword = trim(Request.Form("pass"))

Select case Mode_Type
	case "OpenTopic"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_STATUS = 1 "
				strSQL = strSql & ", T_ARCHIVE_FLAG = 1"
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write "<p class=""c""><span class=""dff hfs""><b>Topic Un-Locked!</b></span></p>" & strLE & _
					"<script language=""javascript1.2"">self.opener.location.reload();</script>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Topic</b></span><br>" & _
					"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
			end if
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Topic</b></span><br>" & _
				"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
		end if
	case "OpenForum"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(chkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "FORUM "
				strSql = strSql & " SET F_STATUS = 1 "
				strSql = strSql & " WHERE FORUM_ID = " & Forum_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write "<p class=""c""><span class=""dff hfs""><b>Forum Un-Locked!</b></span></p>" & strLE & _
					"<script language=""javascript1.2"">self.opener.location.reload();</script>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Forum</b></span><br>" & _
					"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
			end if
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Forum</b></span><br>" & _
				"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
		end if
	case "OpenCategory"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if mLev = 4 then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "CATEGORY "
				strSql = strSql & " SET CAT_STATUS = 1 "
				strSql = strSql & " WHERE CAT_ID = " & Cat_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write "<p class=""c""><span class=""dff hfs""><b>Category Un-Locked!</b></span></p>" & strLE & _
					"<script language=""javascript1.2"">self.opener.location.reload();</script>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Category</b></span><br>" & _
					"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
			end if
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock Category</b></span><br>" & _
				"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
		end if
	case "UnLockMember"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (mLev = 4) and (cLng(chkCanOpen(MemberID,Member_ID)) = 1) then
				'## Forum_SQL
				strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
				strSql = strSql & " SET M_STATUS = 1 "
				strSql = strSql & " WHERE MEMBER_ID = " & Member_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write "<p class=""c""><span class=""dff hfs""><b>Member Un-Locked!</b></span></p>" & strLE & _
					"<script language=""javascript1.2"">self.opener.location.reload();</script>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock a Member</b></span><br>" & _
					"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
			end if
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Un-Lock a Member</b></span><br>" & _
				"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
		end if
	case "UnStickyTopic"
		strEncodedPassword = sha256("" & strPassword)
		mLev = cLng(ChkUser(strDBNTFUserName, strEncodedPassword,-1))
		if mLev > 0 then  '## is Member
			if (chkForumModerator(Forum_ID, strDBNTFUserName) = "1") or (mLev = 4) then
				'## Forum_SQL
				strSql = "UPDATE " & strTablePrefix & "TOPICS "
				strSql = strSql & " SET T_STICKY = " & 0
				strSQL = strSql & ", T_ARCHIVE_FLAG = " & 1
				strSql = strSql & " WHERE TOPIC_ID = " & Topic_ID

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				Response.Write "<p class=""c""><span class=""dff hfs""><b>Topic Made Un-Sticky!</b></span></p>" & strLE & _
					"<script language=""javascript1.2"">self.opener.location.reload();</script>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Make Topic Un-Sticky!</b></span><br>" & _
					"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
			end if
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>No Permissions to Make Topic Un-Sticky!</b></span><br>" & _
				"<br><span class=""dff dfs""><a href=""JavaScript:onClick=history.go(-1)"">Go Back to Re-Authenticate</a></span></p>" & strLE
		end if
	case else
		Response.Write "<p class=""c""><span class=""dff hfs"">"
		select case Mode_Type
			case "Topic"
				Response.Write("Un-Lock Topic")
			case "Forum"
				Response.Write("Un-Lock Forum")
			case "Category"
				Response.Write("Un-Lock Category")
			case "Member"
				Response.Write("Un-Lock Member")
			case "STopic"
				Response.Write("Make Topic Un-Sticky")
		end select
		Response.Write "</span></p>" & strLE & _
				"<p><span class=""dff dfs""><b><span class=""hlfc"">NOTE:&nbsp;</span></b>"
		select case Mode_Type
			case "STopic"
				Response.Write("Only Moderators and Administrators<br>can make a Topic Un-Sticky.")
			case "Member"
				Response.Write("Only Administrators can un-lock a Member.")
			case "Category"
				Response.Write("Only Administrators can un-lock a Category.")
			case "Forum"
				Response.Write("Only Administrators can un-lock a Forum.")
			case "Topic"
				Response.Write("Only Moderators and Administrators<br>can un-lock a Topic.")
		end select
		Response.Write("</span></p>" & strLE)

		Response.Write "<form action=""pop_open.asp?mode="
		select case Mode_Type
			case "Topic"
				Response.Write("OpenTopic")
			case "Forum"
				Response.Write("OpenForum")
			case "Category"
				Response.Write("OpenCategory")
			case "Member"
				Response.Write("UnLockMember")
			case "STopic"
				Response.Write("UnStickyTopic")
		end select
		Response.Write """ method=""Post"">" & strLE & _
			"<input type=""hidden"" name=""TOPIC_ID"" value=""" & Topic_ID & """>" & strLE & _
			"<input type=""hidden"" name=""FORUM_ID"" value=""" & Forum_ID & """>" & strLE & _
			"<input type=""hidden"" name=""CAT_ID"" value=""" & Cat_ID & """>" & strLE & _
			"<input type=""hidden"" name=""MEMBER_ID"" value=""" & Member_ID & """>" & strLE & _
			"<table width=""75%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""pubc"">" & strLE & _
			"<table width=""100%"" cellspacing=""1"" cellpadding=""1"">" & strLE
		if strAuthType = "db" then
		Response.Write "<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">User Name:</span></b></td>" & strLE & _
			"<td class=""putc""><input type=""Text"" maxLength=""25"" name=""Name"" value=""" & chkString(strDBNTUserName,"display") & """ style=""width:150px;""></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">Password:</span></b></td>" & strLE & _
			"<td class=""putc""><input type=""Password"" maxLength=""25"" name=""Pass"" value="""" style=""width:150px;""></td>" & strLE & _
			"</tr>" & strLE
		else
			if strAuthType="nt" then
				Response.Write "<tr>" & strLE & _
					"<td class=""putc nw r""><b><span class=""dff dfs"">NT Account:</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs"">" & chkString(strDBNTUserName,"display") & "</span></td>" & strLE & _
					"</tr>" & strLE
			end if
		end if
		Response.Write "<tr>" & strLE & _
			"<td class=""putc c"" colspan=""2""><Input type=""Submit"" value=""Send""></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE
end select
Call WriteFooterShort
Response.End

function chkCanOpen(fAM_ID, fM_ID)
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_LEVEL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	StrSql = strSql & " WHERE MEMBER_ID = " & fM_ID

	set rsCheck = my_Conn.Execute (strSql)

	if rsCheck.BOF or rsCheck.EOF then
		chkCanOpen = 0 '## No Members Found
	else
		if cLng(rsCheck("MEMBER_ID")) = cLng(fAM_ID) then
			chkCanOpen = 0 '## Can't unlock self
		else
			Select case cLng(rsCheck("M_LEVEL"))
				case 1
					chkCanOpen = 1 '## Can unlock Normal User
				case 2
					chkCanOpen = 1 '## Can unlock Moderator
				case 3
					if fAM_ID <> intAdminMemberID then
						chkCanOpen = 0 '## Only the Forum Admin can unlock other Administrators
					else
						chkCanOpen = 1 '## Forum Admin is ok to unlock other Administrators
					end if
				case else
					chkCanOpen = 0 '## Member doesn't have a Member Level?
			End Select
		end if
	end if

	rsCheck.close
	set rsCheck = nothing
end function
%>
