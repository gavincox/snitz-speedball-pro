<%
'##################################################################################################
'## Snitz Forums 2000 v3.4.07
'##################################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'##		   Huw Reddick and Richard Kinser
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
'##################################################################################################
%>
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_secure.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_profile.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<!--#INCLUDE FILE="inc_moderation.asp" -->
<%
Dim strURLError

if Instr(1,Request.Form("refer"),"search.asp",1) > 0 then
	strRefer = "search.asp"
elseif Instr(1,Request.Form("refer"),"register.asp",1) > 0 then
	strRefer = "default.asp"
else
	strRefer = Request.Form("refer")
end if
if strRefer = "" then strRefer = "default.asp"

if Request.QueryString("id") <> "" and IsNumeric(Request.QueryString("id")) = true then
	ppMember_ID = cLng(Request.QueryString("id"))
else
	ppMember_ID = 0
end if

if strAuthType = "nt" then
	if ChkAccountReg() <> "1" then
		Response.Write "<p class=""c""><span class=""dff dfs hlfc"">" & strLE & _
				"<b>Note:</b> This NT account has not been registered yet, thus the profile is not available.<br>" & strLE
		if strProhibitNewMembers <> "1" then
			Response.Write "If this is your account, <a href=""policy.asp"">click here</a> to register.</span></p>" & strLE
		else
			Response.Write "</span></p>" & strLE
		end if
 		WriteFooter
		Response.End
	end if
end if

select case Request.QueryString("mode")

	case "display" '## Display Profile

		if strDBNTUserName = "" then
			Err_Msg = "You must be logged in to view a Member's Profile"

			Response.Write "<table width=""100%"">" & strLE & _
				"	<tr>" & strLE & _
				"		<td><span class=""dff dfs"">" & strLE & _
				getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
				getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
				getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br>" & strLE & _
				"		" & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member's Profile</span></td>" & strLE & _
				"	</tr>" & strLE & _
				"</table>" & strLE & _
				"	<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem!</span></p>" & strLE & _
				"	<p class=""c""><span class=""dff dfs hlfc"">" & Err_Msg & "</span></p>" & strLE & _
				"	<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Back to Forum</a></span></p>" & strLE & _
				"	<br>" & strLE
				WriteFooterShort
				Response.End
		end if

		'## Forum_SQL
		strSql = "SELECT " & strMemberTablePrefix & "MEMBERS_PENDING.MEMBER_ID"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_NAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_USERNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_FIRSTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LASTNAME"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_TITLE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_PASSWORD"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_AIM"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_ICQ"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_MSN"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_YAHOO"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_COUNTRY"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_POSTS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_CITY"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_STATE"
'		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HIDE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_RECEIVE_EMAIL"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_DATE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_PHOTO_URL"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HOMEPAGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LINK1"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LINK2"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_AGE"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_DOB"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_MARSTATUS"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_SEX"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_OCCUPATION"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_HOBBIES"
		strsql = strsql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_QUOTE"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_LNEWS"
		strSql = strSql & ", " & strMemberTablePrefix & "MEMBERS_PENDING.M_BIO"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " WHERE MEMBER_ID=" & ppMember_ID

		set rs = my_Conn.Execute(strSql)

		if rs.BOF or rs.EOF then
			Err_Msg = "Invalid Member ID!"

			Response.Write "<table width=""100%"">" & strLE & _
				"	<tr>" & strLE & _
				"		<td><span class=""dff dfs"">" & strLE & _
				getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
				getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
				getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br>" & strLE & _
				"		" & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Member's Profile</span></td>" & strLE & _
				"	</tr>" & strLE & _
				"</table>" & strLE & _
				"	<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem!</span></p>" & strLE & _
				"	<p class=""c""><span class=""dff dfs hlfc"">" & Err_Msg & "</span></p>" & strLE & _
				"	<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Back to Forum</a></span></p>" & strLE & _
				"	<br>" & strLE
				WriteFooter
				Response.End
		else
			strMyHobbies = rs("M_HOBBIES")
			strMyQuote = rs("M_QUOTE")
			strMyLNews = rs("M_LNEWS")
			strMyBio = rs("M_BIO")

			intTotalMemberPosts = rs("M_POSTS")
			if intTotalMemberPosts > 0 then
				strMemberDays = DateDiff("d", strToDate(rs("M_DATE")), strToDate(strForumTimeAdjust))
				if strMemberDays = 0 then strMemberDays = 1
				strMemberPostsperDay = round(intTotalMemberPosts/strMemberDays,2)
				if strMemberPostsperDay = 1 then
					strPosts = " post"
				else
					strPosts = " posts"
				end if
			end if

			if strUseExtendedProfile then
				strColspan = " colspan=""2"""
				strIMURL1 = "javascript:openWindow('"
				strIMURL2 = "')"
			else
				strColspan = ""
				strIMURL1 = ""
				strIMURL2 = ""
			end if

			if strUseExtendedProfile then
				Response.Write "<table width=""100%"">" & strLE & _
					"	<tr>" & strLE & _
					"		<td><span class=""dff dfs"">" & strLE & _
					getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
					getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
					getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_accounts_pending.asp"">Members&nbsp;Pending</a><br>" & strLE & _
					"		" & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;" & chkString(rs("M_NAME"),"display") & "'s Profile</span></td>" & strLE & _
					"	</tr>" & strLE & _
					"</table>" & strLE
			end if
			Response.Write "<table width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
				"	<tr>" & strLE & _
				"		<td class=""pb c"" " & strColspan & ">" & strLE & _
				"		<span class=""dff hfs"">Pending Profile<br></span></td>" & strLE & _
				"	</tr>" & strLE & _
				"	<tr>" & strLE & _
				"		<td class=""pb c"" " & strColspan & ">" & strLE & _
				"<table class=""tc"" width=""90%"" cellspacing=""0"" cellpadding=""4"">" & strLE & _
				"	<tr>" & strLE
			if mLev = 4 then
				Response.Write "		<td class=""hcc vat l"">&nbsp;<span class=""dff dfs hfc""><b>" & ChkString(rs("M_NAME"),"display") & "</b></span></td>" & strLE
			else
				Response.Write "		<td class=""hcc vat l""><span class=""dff dfs hfc""><b>&nbsp;" & ChkString(rs("M_NAME"),"display") & "</b></span></td>" & strLE
			end if
			Response.Write "		<td class=""hcc vat r""><span class=""dff dfs hfc"">Pending Since:&nbsp;" & ChkDate(rs("M_DATE"),"",True) & "&nbsp;(" & DateDiff("d",  StrToDate(rs("M_DATE")),  strForumTimeAdjust) & " days)</span></td>" & strLE & _
				"	</tr>" & strLE & _
				"</table>" & strLE & _
				"		</td>" & strLE & _
				"	</tr>" & strLE & _
				"	<tr>" & strLE & _
				"		<td class=""vat pb l"">" & strLE & _
				"<table class=""tc"" width=""90%"" cellspacing=""1"" cellpadding=""0"">" & strLE & _
				"	<tr>" & strLE
			if strUseExtendedProfile then
				Response.Write "		<td class=""vat pb""> width=""35%"" " & strLE & _
					"<table width=""100%"" cellspacing=""0"" cellpadding=""3"">" & strLE
				if trim(rs("M_PHOTO_URL")) = "" or lcase(rs("M_PHOTO_URL")) = "http://" then strPicture = 0
				if strPicture = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">&nbsp;My Picture&nbsp;</span></b></td>" & strLE & _
						"	</tr>" & strLE & _
						"	<tr>" & strLE & _
						"		<td class=""putc c"" colspan=""2"">"
					if Trim(rs("M_PHOTO_URL")) <> "" and lcase(rs("M_PHOTO_URL")) <> "http://" then
						Response.Write "<a href=""" & ChkString(rs("M_PHOTO_URL"), "displayimage") & """>" & getCurrentIcon(ChkString(rs("M_PHOTO_URL"), "displayimage") & "|150|150",ChkString(rs("M_NAME"),"display"),"hspace=""2"" vspace=""2""") & "</a><br><span class=""dff dfs"">Click image for full picture</span>"
					else
						Response.Write getCurrentIcon(strIconPhotoNone,"No Photo Available","hspace=""2"" vspace=""2""")
					end if
					Response.Write "		</td>" & strLE & _
						"	</tr>" & strLE
				end if ' strPicture
				Response.Write "	<tr>" & strLE & _
					"		<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">&nbsp;My Contact Info&nbsp;</span></b></td>" & strLE & _
					"	</tr>" & strLE
				strContacts = 0
				if mLev > 2 or rs("M_RECEIVE_EMAIL") = "1" then
					strContacts = strContacts + 1
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">E-mail Address:&nbsp;</span></b></td>" & strLE
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write "		<td class=""putc nw""><span class=""dff dfs"">" & trim(rs("M_EMAIL")) & "</span></td>" & strLE
					else
						Response.Write "		<td class=""putc""><span class=""dff dfs"">No address specified...</span></td>" & strLE
					end if
					Response.Write "	</tr>" & strLE
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then
					strContacts = strContacts + 1
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">AIM:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon(strIconAIM,"","class=""vam""") & "&nbsp;<a href=""" & strIMURL1 & "pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then
					strContacts = strContacts + 1
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">ICQ:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","class=""vam""") & "&nbsp;<a href=""" & strIMURL1 & "pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & strIMURL2 & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then
					strContacts = strContacts + 1
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = partss(1)

					Response.Write "		<script type=""text/javascript"">" & strLE & _
						"			function MSNjs() {" & strLE & _
						"			var tag1 = '" & strtag1 & "';" & strLE & _
						"			var tag2 = '" & strtag2 & "';" & strLE & _
						"			var tag3 = '" & strtag3 & "';" & strLE & _
						"			document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & strLE & _
						"		</script>" & strLE
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">MSN:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon(strIconMSNM,"","class=""vam""") & "&nbsp;<script type=""text/javascript"">MSNjs()</script>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then
					strContacts = strContacts + 1
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">YAHOO IM:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs""><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strContacts = 0 then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw c"" colspan=""2""><span class=""dff dfs"">No info specified...</span></td>" & strLE & _
						"	</tr>" & strLE
				end if


				if (strHomepage + strFavLinks) > 0 then

					Response.Write "	<tr>" & strLE & _
						"		<td class=""ccc c"" colspan=""2"">" & strLE & _
						"		<b><span class=""dff dfs cfc"">Links&nbsp;</span></b></td>" & strLE
					if strHomepage = "1" then
						Response.Write "	<tr>" & strLE & _
							"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">Homepage:&nbsp;</span></b></td>" & strLE
						if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
							Response.Write "		<td class=""putc""><span class=""dff dfs""><a href=""" & rs("M_HOMEPAGE") & """ target=""_blank"">" & rs("M_HOMEPAGE") & "</a>&nbsp;</span></td>" & strLE
						else
							Response.Write "		<td class=""putc""><span class=""dff dfs"">No homepage specified...</span></td>" & strLE
						end if
						Response.Write "	</tr>" & strLE
					end if
					if strFavLinks = "1" then
						Response.Write "	<tr>" & strLE & _
							"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">Cool Links:&nbsp;</span></b></td>" & strLE
						if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then
							Response.Write "		<td class=""putc""><span class=""dff dfs""><a href=""" & rs("M_LINK1") & """ target=""_blank"">" & rs("M_LINK1") & "</a>&nbsp;</span></td>" & strLE
							if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
								Response.Write "	</tr>" & strLE & _
									"	<tr>" & strLE & _
									"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">&nbsp;</span></b></td>" & strLE & _
									"		<td class=""putc""><span class=""dff dfs""><a href=""" & rs("M_LINK2") & """ target=""_blank"">" & rs("M_LINK2") & "</a>&nbsp;</span></td>" & strLE
							end if
						else
							Response.Write "		<td class=""putc""><span class=""dff dfs"">No link specified...</span></td>" & strLE
						end if
						Response.Write "	</tr>" & strLE
					end if
				end if ' strRecentTopics
				Response.Write "</table>" & strLE & _
					"		</td>" & strLE & _
					"		<td class=""vat pb"" width=""3%"">&nbsp;</td>" & strLE
			end if ' UseExtendedMemberProfile
			Response.Write "		<td class=""vat pb"">" & strLE & _
				"<table class=""vat"" width=""100%"" cellspacing=""0"" cellpadding=""3"">" & strLE & _
				"	<tr>" & strLE & _
				"		<td class=""ccc vat c"" colspan=""2""><b><span class=""dff dfs cfc"">Basics</span></b></td>" & strLE & _
				"	</tr>" & strLE & _
				"	<tr>" & strLE & _
				"		<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">User Name:&nbsp;</span></b></td>" & strLE & _
				"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_NAME"),"display") & "&nbsp;</span></td>" & strLE & _
				"	</tr>" & strLE
			if strAuthType = "nt" then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Your Account:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_USERNAME"),"display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if strFullName = "1" and (Trim(rs("M_FIRSTNAME")) <> "" or Trim(rs("M_LASTNAME")) <> "" ) then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Real Name:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_FIRSTNAME"), "display") & "&nbsp;" & ChkString(rs("M_LASTNAME"), "display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if (strCity = "1" and Trim(rs("M_CITY")) <> "") or (strCountry = "1" and Trim(rs("M_COUNTRY")) <> "") or (strState = "1" and Trim(rs("M_STATE")) <> "") then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Location:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">"
				myCity = ChkString(rs("M_CITY"),"display")
				myState = ChkString(rs("M_STATE"),"display")
				myCountry = ChkString(rs("M_COUNTRY"),"display")
				myLocation = ""

				if myCity <> "" and myCity <> " " then
					myLocation = myCity
				end if

				if myLocation <> "" then
					if myState <> "" and myState <> " " then
						myLocation = myLocation & ",&nbsp;" & myState
					end if
				else
					if myState <> "" and myState <> " " then
						myLocation = myState
					end if
				end if

				if myLocation <> "" then
					if myCountry <> "" and myCountry <> " " then
						myLocation = myLocation & "<br>" & myCountry
					end if
				else
					if myCountry <> "" and myCountry <> " " then
						myLocation = myCountry
					end if
				end if
				Response.Write myLocation
				Response.Write "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if (strAge = "1" and Trim(rs("M_AGE")) <> "") then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Age:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_AGE"), "display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			strDOB = rs("M_DOB")
			if (strAgeDOB = "1" and Trim(strDOB) <> "") then
			strDOB = DOBToDate(strDOB)
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Age:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & DisplayUsersAge(strDOB) & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if (strMarStatus = "1" and Trim(rs("M_MARSTATUS")) <> "") then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Marital Status:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_MARSTATUS"), "display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if (strSex = "1" and Trim(rs("M_SEX")) <> "") then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Gender:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_SEX"), "display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if (strOccupation = "1" and Trim(rs("M_OCCUPATION")) <> "") then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Occupation:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(rs("M_OCCUPATION"), "display") & "</span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if intTotalMemberPosts > 0 then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""putc vat nw r""><b><span class=""dff dfs"">Total Posts:&nbsp;</span></b></td>" & strLE & _
					"		<td class=""putc""><span class=""dff dfs"">" & ChkString(intTotalMemberPosts, "display") & "<br><span class=""ffs"">[" & strMemberPostsperDay & strPosts & " per day]<br><a href=""search.asp?mode=DoIt&MEMBER_ID=" & rs("MEMBER_ID") & """>Find all non-archived posts by " & chkString(rs("M_NAME"),"display") & "</a></span></span></td>" & strLE & _
					"	</tr>" & strLE
			end if
			if not(strUseExtendedProfile) then
				if rs("M_RECEIVE_EMAIL") = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">E-mail Address:&nbsp;</span></b></td>" & strLE
					if Trim(rs("M_EMAIL")) <> "" then
						Response.Write "		<td class=""putc nw""><span class=""dff dfs""><a href=""pop_mail.asp?id=" & rs("MEMBER_ID") & """>Click to send an E-Mail</a>&nbsp;</span></td>" & strLE
					else
						Response.Write "		<td class=""putc""><span class=""dff dfs"">No address specified...</span></td>" & strLE
					end if
					Response.Write "	</tr>" & strLE
				end if
				if strAIM = "1" and Trim(rs("M_AIM")) <> "" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">AIM:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon(strIconAIM,"","class=""vam""") & "&nbsp;<a href=""pop_messengers.asp?mode=AIM&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_AIM"), "display") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strICQ = "1" and Trim(rs("M_ICQ")) <> "" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">ICQ:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon("http://online.mirabilis.com/scripts/online.dll?icq=" & ChkString(rs("M_ICQ"), "urlpath") & "&img=5|18|18","","class=""vam""") & "&nbsp;<a href=""pop_messengers.asp?mode=ICQ&ID=" & rs("MEMBER_ID") & """>" & ChkString(rs("M_ICQ"), "display") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strMSN = "1" and Trim(rs("M_MSN")) <> "" then
					parts = split(rs("M_MSN"),"@")
					strtag1 = parts(0)
					partss = split(parts(1),".")
					strtag2 = partss(0)
					strtag3 = partss(1)

					Response.Write "<script type=""text/javascript"">" & strLE & _
						"	function MSNjs() {" & strLE & _
						"		var tag1 = '" & strtag1 & "';" & strLE & _
						"		var tag2 = '" & strtag2 & "';" & strLE & _
						"		var tag3 = '" & strtag3 & "';" & strLE & _
						"		document.write(tag1 + ""@"" + tag2 + ""."" + tag3) }" & strLE & _
						"</script>" & strLE

					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">MSN:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs"">" & getCurrentIcon(strIconMSNM,"","class=""vam""") & "&nbsp;<script type=""text/javascript"">MSNjs()</script>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strYAHOO = "1" and Trim(rs("M_YAHOO")) <> "" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r""><b><span class=""dff dfs"">YAHOO IM:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc""><span class=""dff dfs""><a href=""http://edit.yahoo.com/config/send_webmesg?.target=" & ChkString(rs("M_YAHOO"), "urlpath") & "&.src=pg"" target=""_blank"">" & getCurrentIcon("http://opi.yahoo.com/online?u=" & ChkString(rs("M_YAHOO"), "urlpath") & "&m=g&t=2|125|25","","") & "</a>&nbsp;</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
			end if
			if IsNull(strMyBio) or trim(strMyBio) = "" then strBio = 0
			if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then strHobbies = 0
			if IsNull(strMyLNews) or trim(strMyLNews) = "" then strLNews = 0
			if IsNull(strMyQuote) or trim(strMyQuote) = "" then strQuote = 0
			if (strBio + strHobbies + strLNews + strQuote) > 0 then
				Response.Write "	<tr>" & strLE & _
					"		<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">More About Me</span></b></td>" & strLE & _
					"	</tr>" & strLE
				if strBio = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Bio:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc vat""><span class=""dff dfs"">"
					if IsNull(strMyBio) or trim(strMyBio) = "" then Response.Write("-") else Response.Write(formatStr(strMyBio))
					Response.Write "</span></td>" & strLE & _
						"		</tr>" & strLE
				end if
				if strHobbies = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Hobbies:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc vat""><span class=""dff dfs"">"
					if IsNull(strMyHobbies) or trim(strMyHobbies) = "" then Response.Write("-") else Response.Write(formatStr(strMyHobbies))
					Response.Write "</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strLNews = "1" then
					Response.Write "		<tr>" & strLE & _
						"		<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Latest News:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc vat""><span class=""dff dfs"">"
					if IsNull(strMyLNews) or trim(strMyLNews) = "" then Response.Write("-") else Response.Write(formatStr(strMyLNews))
					Response.Write "</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strQuote = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Favorite Quote:&nbsp;</span></b></td>" & strLE & _
						"		<td class=""putc vat""><span class=""dff dfs"">"
					if IsNull(strMyQuote) or Trim(strMyQuote) = "" then Response.Write("-") else Response.Write(formatStr(strMyQuote))
					Response.Write "</span></td>" & strLE & _
						"	</tr>" & strLE
				end if
			end if
			if (strHomepage + strFavLinks) > 0 and not(strRecentTopics = "0" and strUseExtendedProfile) then
				if strUseExtendedProfile then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">Links&nbsp;</span></b></td>" & strLE & _
						"	</tr>" & strLE
				end if
				if strHomepage = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">Homepage:&nbsp;</span></b></td>" & strLE
					if Trim(rs("M_HOMEPAGE")) <> "" and lcase(trim(rs("M_HOMEPAGE"))) <> "http://" and Trim(lcase(rs("M_HOMEPAGE"))) <> "https://" then
						Response.Write "		<td class=""putc""><span class=""dff dfs""><a href=""" & ChkString(rs("M_HOMEPAGE"), "display") & """ target=""_blank"">" & ChkString(rs("M_HOMEPAGE"), "display") & "</a>&nbsp;</span></td>" & strLE
					else
						Response.Write "		<td class=""putc""><span class=""dff dfs"">No homepage specified...</span></td>" & strLE
					end if
					Response.Write "	</tr>" & strLE
				end if
				if strFavLinks = "1" then
					Response.Write "	<tr>" & strLE & _
						"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">Cool Links:&nbsp;</span></b></td>" & strLE
					if Trim(rs("M_LINK1")) <> "" and lcase(trim(rs("M_LINK1"))) <> "http://" and Trim(lcase(rs("M_LINK1"))) <> "https://" then
						Response.Write "		<td class=""putc""><span class=""dff dfs""><a href=""" & ChkString(rs("M_LINK1"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK1"), "display") & "</a>&nbsp;</span></td>" & strLE
						if Trim(rs("M_LINK2")) <> "" and lcase(trim(rs("M_LINK2"))) <> "http://" and Trim(lcase(rs("M_LINK2"))) <> "https://" then
							Response.Write "	</tr>" & strLE & _
								"	<tr>" & strLE & _
								"		<td class=""putc nw r"" width=""10%""><b><span class=""dff dfs"">&nbsp;</span></b></td>" & strLE & _
								"		<td class=""putc""><span class=""dff dfs""><a href=""" & ChkString(rs("M_LINK2"), "display") & """ target=""_blank"">" & ChkString(rs("M_LINK2"), "display") & "</a>&nbsp;</span></td>" & strLE
						end if
					else
						Response.Write "								<td class=""putc""><span class=""dff dfs"">No link specified...</span></td>" & strLE
					end if
					Response.Write "							</tr>" & strLE
				end if
			end if
			Response.Write "							</table>" & strLE & _
					"								</td>" & strLE & _
					"							</tr>" & strLE & _
					"						</table>" & strLE & _
					"					</td>" & strLE & _
					"</tr>" & strLE & _
					"</table><br>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE
				Response.Write "<tr>" & strLE & _
						"<td class=""pb nw c"">" & strLE
		end if
	case else
		Response.Redirect("default.asp")
end select

set rs = nothing
	WriteFooter

Function IsValidURL(sValidate)
	Dim sInvalidChars
	Dim bTemp
	Dim i

	if trim(sValidate) = "" then IsValidURL = true : exit function
	sInvalidChars = """;+()*'<>"
	for i = 1 To Len(sInvalidChars)
		if InStr(sValidate, Mid(sInvalidChars, i, 1)) > 0 then bTemp = True
		if bTemp then strURLError = "<br>&bull;&nbsp;cannot contain any of the following characters:  "" ; + ( ) * ' < > "
		if bTemp then Exit For
	next
	if not bTemp then
		for i = 1 to Len(sValidate)
			if Asc(Mid(sValidate, i, 1)) = 160 then bTemp = True
			if bTemp then strURLError = "<br>&bull;&nbsp;cannot contain any spaces "
			if bTemp then Exit For
		next
	end if

	' extra checks
	' check to make sure URL begins with http:// or https://
	if not bTemp then
		bTemp = (lcase(left(sValidate, 7)) <> "http://") and (lcase(left(sValidate, 8)) <> "https://")
		if bTemp then strURLError = "<br>&bull;&nbsp;must begin with either http:// or https:// "
	end if
	' check to make sure URL is 255 characters or less
	if not bTemp then
		bTemp = len(sValidate) > 255
		if bTemp then strURLError = "<br>&bull;&nbsp;cannot be more than 255 characters "
	end if
	' no two consecutive dots
	if not bTemp then
		bTemp = InStr(sValidate, "..") > 0
		if bTemp then strURLError = "<br>&bull;&nbsp;cannot contain consecutive periods "
	end if
	'no spaces
	if not bTemp then
		bTemp = InStr(sValidate, " ") > 0
		if bTemp then strURLError = "<br>&bull;&nbsp;cannot contain any spaces "
	end if
	if not bTemp then
		bTemp = (len(sValidate) <> len(Trim(sValidate)))
		if bTemp then strURLError = "<br>&bull;&nbsp;cannot contain any spaces "
	end if 'Addition for leading and trailing spaces

	' if any of the above are true, invalid string
	IsValidURL = Not bTemp
End Function
%>