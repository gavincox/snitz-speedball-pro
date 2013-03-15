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
Sub DisplayProfileForm
	on error resume next
	strMode = Request.QueryString("mode")

	Response.Write "<table class=""vat tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""pb c"" " & strColSpan & "><p><span class=""dff dfs""><b>All Fields marked with <span class=""hfs hlfc"">*</span> are required</b>"
	if lcase(strEmail) = "1" and strEmailVal = "1" then
		if strMode = "Register" then
			Response.Write("<br><small>To complete your registration, you need to have a valid e-mail address.</small>")
		else
			if strMode <> "goModify" then
				Response.Write("<br><small>If you change your e-mail address, a confirmation e-mail will be sent to your new address.<br>Please make sure it is a valid address.</small>")
			else
				Response.Write("<br><small>If you change the e-mail address, a confirmation e-mail will be sent to the new address.<br>Please make sure it is a valid address.</small>")
			end if
		end if
	end if
	Response.Write "</span></p></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat pb l"">" & strLE & _
			"<table class=""tc"" width=""80%"" cellspacing=""1"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE

	if strUseExtendedProfile then
		Response.Write "<td class=""putc vat"" width=""50%"">" & strLE & _
				"<table width=""100%"" cellspacing=""0"" cellpadding=""1"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">&nbsp;Contact Info&nbsp;</span></b></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> E-mail Address:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs""><input name=""Email"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then  Response.Write(rs("M_EMAIL"))
		Response.Write """>" & strLE & _
				"<input type=""hidden"" name=""Email2"" value="""
		if strMode <> "Register" then Response.Write(rs("M_EMAIL"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
		if strMode = "Register" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> E-mail Address Again:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Email3"" size=""25"" maxLength=""50"" value=""""></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		Response.Write "<tr class=""vam"">" & strLE & _
				"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff ffs"">Allow Forum Members<br>to Send you E-Mail?:&nbsp;</span></b></td>" & strLE
		if strMode = "Register" then
			Response.Write "<td class=""putc vam""><span class=""dff dfs"">" & strLE & _
					"<select name=""ReceiveEMail"">" & strLE & _
					"                      	<option value=""1"" selected>Yes</option>" & strLE & _
					"                      	<option value=""0"">No</option>" & strLE & _
					"</select></span></td>" & strLE
		else
			Response.Write "<td class=""putc vam""><span class=""dff ffs"">" & strLE & _
					"<select name=""ReceiveEMail"">" & strLE & _
					"                      	<option value=""1"""
			if rs("M_RECEIVE_EMAIL") <> "0" then Response.Write(" selected")
			Response.Write ">Yes</option>" & strLE & _
					"                      	<option value=""0"""
			if rs("M_RECEIVE_EMAIL") = "0" then Response.Write(" selected")
			Response.Write ">No</option>" & strLE & _
					"</select></span></td>" & strLE
		end if
		Response.Write "</tr>" & strLE
		if strMode = "goModify" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff dfs"">Initial IP:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><a href=""" & strIPLookup & ChkString(rs("M_IP"), "display") & """ target=""_blank"">" & ChkString(rs("M_IP"), "display") & "</a></span></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff dfs"">Last IP:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><a href=""" & strIPLookup & ChkString(rs("M_LAST_IP"), "display") & """ target=""_blank"">" & ChkString(rs("M_LAST_IP"), "display") & "</a></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strAIM = "1" then
			Response.Write "<tr>" & strLE
			if strReqAIM = "1" then
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> AIM: </span></b></td>" & strLE
			else
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">AIM: </span></b></td>" & strLE
			end if
			Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""AIM"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_AIM"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strICQ = "1" then
			Response.Write "<tr>" & strLE
			if strReqICQ = "1" then
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> ICQ: </span></b></td>" & strLE
			else
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">ICQ: </span></b></td>" & strLE
			end if
			Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""ICQ"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_ICQ"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strMSN = "1" then
			Response.Write "<tr>" & strLE
			if strReqMSN = "1" then
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> MSN: </span></b></td>" & strLE
			else
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">MSN: </span></b></td>" & strLE
			end if
			Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""MSN"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_MSN"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strYAHOO = "1" then
			Response.Write "<tr>" & strLE
			if strReqYAHOO = "1" then
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Yahoo!: </span></b></td>" & strLE
			else
				Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Yahoo!: </span></b></td>" & strLE
			end if
			Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""YAHOO"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_YAHOO"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if (strHomepage + strFavLinks) > 0 then
			Response.Write "<tr>" & strLE & _
					"<td class=""ccc c"" colspan=""2"">" & strLE & _
					"<b><span class=""dff dfs cfc"">Links&nbsp;</span></b></td>" & strLE & _
					"</tr>" & strLE
			if strHomepage = "1" then
				Response.Write "<tr>" & strLE
				if strReqHomepage = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Homepage: </span></b></td>" & strLE & _
							"<td class=""putc""><span class=""dff dfs""><input name=""Homepage"" size=""25"" maxLength=""255"" value="""
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Homepage: </span></b></td>" & strLE & _
							"<td class=""putc""><span class=""dff dfs""><input name=""Homepage"" size=""25"" maxLength=""255"" value="""
				end if
				if strMode <> "Register" then
					if ChkString(rs("M_HOMEPAGE"), "display") <> " " and lcase(rs("M_HOMEPAGE")) <> "http://" then Response.Write(rs("M_HOMEPAGE")) else Response.Write("http://") end if
				else
					Response.Write("http://")
				end if
				Response.Write """></span></td>" & strLE & _
						"</tr>" & strLE
			end if
			if strFavLinks = "1" then
				Response.Write "<tr>" & strLE
				if strReqFavLinks = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Cool Links: </span></b></td>" & strLE & _
							"<td class=""putc""><span class=""dff dfs""><input name=""Link1"" size=""25"" maxLength=""255"" value="""
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Cool Links: </span></b></td>" & strLE & _
							"<td class=""putc""><span class=""dff dfs""><input name=""Link1"" size=""25"" maxLength=""255"" value="""
				end if
				if strMode <> "Register" then
					if rs("M_LINK1") <> " " and lcase(rs("M_LINK1")) <> "http://" then Response.Write(ChkString(rs("M_LINK1"), "display")) else Response.Write("http://")
				else
					Response.Write("http://")
				end if
				Response.Write """></span></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">&nbsp;</span></b></td>" & strLE & _
						"<td class=""putc""><span class=""dff dfs""><input name=""Link2"" size=""25"" maxLength=""255"" value="""
				if strMode <> "Register" then
					if rs("M_LINK2") <> " " and lcase(rs("M_LINK2")) <> "http://" then Response.Write(ChkString(rs("M_LINK2"), "display")) else Response.Write("http://")
				else
					Response.Write("http://")
				end if
				Response.Write """></span></td>" & strLE & _
						"</tr>" & strLE
			end if
		end if
		if strPicture = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""ccc c"" colspan=""2"">" & strLE & _
					"<b><span class=""dff dfs cfc"">Picture</span></b></td>" & strLE & _
					"</tr>" & strLE
			if strReqPicture = "1" then
				Response.Write "<tr>" & strLE & _
						"<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Picture URL: </span></b></td>" & strLE & _
						"<td class=""putc""><span class=""dff dfs""><input name=""Photo_URL"" size=""25"" maxLength=""255"" value="""
			else
				Response.Write "<tr>" & strLE & _
							"<td class=""putc vat nw r""><b><span class=""dff dfs"">Picture URL: </span></b></td>" & strLE & _
							"<td class=""putc""><span class=""dff dfs""><input name=""Photo_URL"" size=""25"" maxLength=""255"" value="""
			end if
			if strMode <> "Register" then
				if rs("M_PHOTO_URL") <> " " and lcase(rs("M_PHOTO_URL")) <> "http://" then Response.Write(ChkString(rs("M_PHOTO_URL"), "displayimage")) else Response.Write("http://")
			else
				Response.Write("http://")
			end if
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if (strBio + strHobbies + strLNews + strQuote)	> 0 then
			if strMode <> "Register" then
				strMyHobbies = rs("M_HOBBIES")
				strMyLNews = rs("M_LNEWS")
				strMyQuote = rs("M_QUOTE")
				strMyBio = rs("M_BIO")
			else
				strMyHobbies = ""
				strMyLNews = ""
				strMyQuote = ""
				strMyBio = ""
			end if
			Response.Write "<tr>" & strLE & _
					"<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">More About Me</span></b></td>" & strLE & _
					"</tr>" & strLE
			if strHobbies = "1" then
				Response.Write "<tr>" & strLE
				if strReqHobbies = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Hobbies: </span></b></td>" & strLE
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Hobbies: </span></b></td>" & strLE
				end if
				Response.Write "<td class=""putc vat""><span class=""dff dfs""><textarea name=""Hobbies"" cols=""30"" rows=""4"">" & Trim(cleancode(strMyHobbies)) & "</textarea></span></td>" & strLE & _
						"</tr>" & strLE
			end if
			if strLNEWS = "1" then
				Response.Write "<tr>" & strLE
				if strReqLNEWS = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Latest News: </span></b></td>" & strLE
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Latest News: </span></b></td>" & strLE
				end if
				Response.Write "<td class=""putc vat""><span class=""dff dfs""><textarea name=""LNews"" cols=""30"" rows=""4"">" & Trim(cleancode(strMyLNews)) & "</textarea></span></td>" & strLE & _
						"</tr>" & strLE
			end if
			if strQuote = "1" then
				Response.Write "<tr>" & strLE
				if strReqQuote = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Favorite Quote: </span></b></td>" & strLE
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Favorite Quote: </span></b></td>" & strLE
				end if
				Response.Write "<td class=""putc vat""><span class=""dff dfs""><textarea name=""Quote"" cols=""30"" rows=""4"">" & Trim(cleancode(strMyQuote)) & "</textarea></span></td>" & strLE & _
						"</tr>" & strLE
			end if
			if strBio = "1" then
				Response.Write "<tr>" & strLE
				if strReqBio = "1" then
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Bio: </span></b></td>" & strLE
				else
					Response.Write "<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Bio: </span></b></td>" & strLE
				end if
				Response.Write "<td class=""putc vat""><span class=""dff dfs""><textarea name=""Bio"" cols=""30"" rows=""4"">" & Trim(cleancode(strMyBio)) & "</textarea></span></td>" & strLE & _
						"</tr>" & strLE
			end if
		end if
		Response.Write "</table>" & strLE & _
				"</td>" & strLE
	end if 'extended profile

	Response.Write "<td class=""putc vat"">" & strLE & _
			"<table width=""100%"" cellspacing=""0"" cellpadding=""1"">" & strLE & _
			"<tr>" & strLE & _
			"<td colspan=""2"" class=""ccc vat c""><b><span class=""dff dfs cfc"">Basics</span></b></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> User Name:&nbsp;</span></b></td>" & strLE & _
			"<td class=""putc""><span class=""dff dfs"">" & strLE
	if (strMode = "goEdit") or (strMode = "goModify" and cLng(Request.Form("MEMBER_ID")) = cLng(intAdminMemberID)) then
		Response.Write "<span class=""dff dfs"">" & ChkString(rs("M_NAME"), "display") & "</span>" & strLE & _
				"<input type=""hidden"" name=""Name"" value=""" & chkString(rs("M_NAME"), "sqlstring") & """>" & strLE
	else
		Response.Write "<input name=""Name"" size=""25"" maxLength=""25"" value="""
		if strMode <> "Register" then Response.Write(ChkString(rs("M_NAME"), "display"))
		Response.Write """>" & strLE
	end if
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE
	if strMode = "goModify" then
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">Title:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs""><input name=""Title"" size=""25"" maxLength=""50"" value=""" & CleanCode(rs("M_TITLE")) & """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strAuthType = "nt" then
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Your Account:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs"">" & strLE
		if Request.Form("Method_Type") = "Modify" then
			Response.Write "<input name=""Account"" value=""" & ChkString(rs("M_USERNAME"), "display") & """>" & strLE
		else
			Response.Write "                      " & Session(strCookieURL & "userid") & "<input type=""hidden"" name=""Account"" value=""" & Session(strCookieURL & "userid") & """>" & strLE
		end if
		Response.Write "</span></td>" & strLE & _
				"</tr>" & strLE
	else
		if strMode = "Register" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Password:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Password"" type=""Password"" size=""25"" maxLength=""25"" value=""""></span></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Password Again:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Password2"" type=""Password"" size=""25"" maxLength=""25"" value=""""></span></td>" & strLE & _
					"</tr>" & strLE
	        else
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff dfs"">&nbsp;New Password:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Password"" type=""Password"" size=""25"" maxLength=""25"" value="""">" & strLE & _
					"<input name=""Password-d"" type=""hidden"" value=""" & rs("M_PASSWORD") & """></span></td>" & strLE & _
					"</tr>" & strLE
			if strMode = "goEdit" then
				Response.Write "<tr>" & strLE & _
						"<td class=""putc vat nw r""><b><span class=""dff dfs"">&nbsp;New Password Again:&nbsp;</span></b></td>" & strLE & _
						"<td class=""putc""><span class=""dff dfs""><input name=""Password2"" type=""Password"" size=""25"" maxLength=""25"" value=""""></span></td>" & strLE & _
						"</tr>" & strLE
			end if
		end if
	end if
	if strFullName = "1" then
		Response.Write "<tr>" & strLE
		If strReqFullName = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> First Name:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">First Name:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""FirstName"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then Response.Write(rs("M_FIRSTNAME"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE
		If strReqFullName = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Last Name:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Last Name:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""LastName"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then Response.Write(rs("M_LASTNAME"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strCity = "1" then
		Response.Write "<tr>" & strLE
		If strReqCity = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> City:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">City:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""City"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then Response.Write(rs("M_CITY"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strState = "1" then
		Response.Write "<tr>" & strLE
		If strReqState = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> State:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">State:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""State"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then Response.Write(rs("M_STATE"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strCountry = "1" then
		Response.Write "<tr>" & strLE
		If strReqCountry = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Country:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Country:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs"">" & strLE & _
				"<select name=""Country"" size=""1"">" & strLE
		if strMode <> "Register" then
			Response.Write("<option selected value=""" & rs("M_COUNTRY") & """>" & ChkString(rs("M_COUNTRY"), "display") & "</option>" & strLE)
		else
			Response.Write "<option value=""""></option>" & strLE
		end if
		Response.Write "<option value="""">None</option>" & strLE
%>
		<!--#INCLUDE FILE="inc_countrylist.asp"-->
<%
		Response.Write "</select></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strMinAge > 0 or strReqAge = "1" or strReqAgeDOB = "1" then
		strReq = "<span class=""hlfc"">*</span>"
	else
		strReq = ""
	end if
	if strAge = "1" then
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">" & strReq & "&nbsp;Age:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs""><input name=""Age"" size=""5"" maxLength=""3"" value="""
		if strMode <> "Register" then Response.Write(ChkString(rs("M_AGE"), "display"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strAgeDOB = "1" then
		strDOByear = ""
		strDOBmonth = ""
		strDOBday = ""
		if strMode <> "Register" then
			strMDOB = trim(ChkString(rs("M_DOB"), "display"))
			if len(strMDOB) > 0 then
				strDOByear = cInt(left(strMDOB, 4))
				strDOBmonth = cInt(mid(strMDOB, 5, 2))
				strDOBday = cInt(right(strMDOB, 2))
			end if
		end if
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">" & strReq & "&nbsp;Birth Date: </span></b></td>" & strLE & _
				"<td class=""putc nw""><span class=""dff dfs"">" & strLE & _
				"<select name=""year"" id=""year"" onchange=""DateSelector(0);"">" & strLE & _
				"<option value=""""" & chkSelect("", strDOByear) & ">Year</option>" & strLE
		intStartYear = cInt(year(strForumTimeAdjust) - strMinAge)
		for intYear = intStartYear to 1900 step -1
			Response.Write "<option value=""" & intYear & """"
			if strMode <> "Register" and len(strMDOB) > 0 then
				Response.Write chkSelect(intYear, strDOByear)
			end if
			Response.Write ">" & intYear & "</option>" & strLE
		next
		Response.Write "</select> "& strLE & _
				"<select name=""month"" id=""month"" style=""visibility:visible;"" onchange=""DateSelector(1);"">"& strLE & _
				"<option value=""""" & chkSelect("", strDOBmonth) & ">Month</option>" & strLE
		for intMonth = 1 to 12
			Response.Write "<option value=""" & doublenum(intMonth) & """"
			if strMode <> "Register" and len(strMDOB) > 0 then
				Response.Write chkSelect(intMonth, strDOBmonth)
			end if
			Response.Write ">" & monthname(intMonth) & "</option>" & strLE
		next
		Response.Write "</select> "& strLE & _
				"<select name=""day"" id=""day"" style=""visibility:visible;"">" & strLE & _
				"<option value=""""" & chkSelect("", strDOBday) & ">Day</option>" & strLE
		for intDay = 1 to 31
			Response.Write "<option value=""" & doublenum(intDay) & """"
			if strMode <> "Register" and len(strMDOB) > 0 then
				Response.Write chkSelect(intDay, strDOBday)
			end if
			Response.Write ">" & intDay & "</option>" & strLE
		next
		Response.Write  "</select></span>"& strLE & _
				"<script type=""text/javascript"" src=""inc_datepicker.js""></script></td>"& strLE & _
				"</tr>" & strLE
	end if
	if strSex = "1" then
		Response.Write "<tr>" & strLE
		If strReqSex = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Gender:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Gender:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs"">" & strLE & _
				"<select name=""Sex"" size=""1"">" & strLE & _
				"<option value="""""
		if strMode <> "Register" then
			if rs("M_SEX") = "" then Response.Write(" selected")
		else
			Response.Write(" selected")
		end if
		Response.Write ">Not specified&nbsp;</option>" & strLE & _
				"<option value=""Male"""
		if strMode <> "Register" then
			if rs("M_SEX") = "Male" then Response.Write(" selected")
		end if
		Response.Write ">Male&nbsp;</option>" & strLE & _
				"<option value=""Female"""
		if strMode <> "Register" then
			if rs("M_SEX") = "Female" then Response.Write(" selected")
		end if
		Response.Write ">Female&nbsp;</option>" & strLE & _
				"</select></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strMarStatus = "1" then
		Response.Write "<tr>" & strLE
		If strReqMarStatus = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Marital Status:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Marital Status:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""MarStatus"" size=""25"" maxLength=""25"" value="""
		if strMode <> "Register" then Response.Write(ChkString(rs("M_MARSTATUS"), "display"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strOccupation = "1" then
		Response.Write "<tr>" & strLE
		If strReqOccupation = "1" Then
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs""><span class=""hlfc"">*</span> Occupation:&nbsp;</span></b></td>" & strLE
		Else
			Response.Write "<td class=""putc vat nw r""><b><span class=""dff dfs"">Occupation:&nbsp;</span></b></td>" & strLE
		End If
		Response.Write "<td class=""putc""><span class=""dff dfs""><input name=""Occupation"" size=""25"" maxLength=""255"" value="""
		if strMode <> "Register" then Response.Write(ChkString(rs("M_OCCUPATION"), "display"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strMode = "goModify" then
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs""># of Posts:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs""><input name=""Posts"" size=""5"" maxLength=""10"" value=""" & ChkString(rs("M_POSTS"), "display") & """></span></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strSignatures = "1" then
		if strMode <> "Register" then
			strTxtSig = rs("M_SIG")
		end if
		Response.Write "<script type=""text/javascript"" src=""inc_code.js""></script>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">Signature:&nbsp;</span></b><br>" & strLE & _
				"<span style=""font-size: 4px;""><br></span>" & strLE & _
				"<table>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw l""><span class=""dff ffs"">" & strLE
		if strAllowHTML = "1" then
			Response.Write "                            * HTML is ON<br>" & strLE
		else
			Response.Write "                            * HTML is OFF<br>" & strLE
		end if
		if strAllowForumCode = "1" then
			Response.Write "                            * <a href=""JavaScript:openWindow6('pop_forum_code.asp')"">Forum Code</a> is ON<br>" & strLE
		else
			Response.Write "                            * Forum Code is OFF<br>" & strLE
		end if
		Response.Write "</span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"<span style=""font-size: 4px;""><br></span>" & strLE & _
				"<input name=""Preview"" type=""button"" value=""Preview"" onclick=""OpenSigPreview()"">&nbsp;</td>" & strLE & _
				"<td class=""putc""><textarea name=""Sig"" cols=""25"" rows=""4"">" & Trim(cleancode(strTxtSig)) & "</textarea></td>" & strLE & _
				"</tr>" & strLE
		if strMode <> "goModify" then
			if strDSignatures = "1" then
				Response.Write "<tr>" & strLE & _
						"<td class=""putc vat nw r""><b><span class=""dff ffs"">View Signatures<br>in Posts?:&nbsp;</span></b></td>" & strLE & _
						"<td class=""putc vam""><span class=""dff dfs"">" & strLE & _
						"<select name=""ViewSig"">" & strLE
				if strMode = "Register" then
					Response.Write "                      	<option value=""1"" selected>Yes</option>" & strLE & _
							"                      	<option value=""0"">No</option>" & strLE
				else
					Response.Write "                      	<option value=""1""" & chkSelect(rs("M_VIEW_SIG"),1) & ">Yes</option>" & strLE & _
							"                      	<option value=""0""" & chkSelect(rs("M_VIEW_SIG"),0) & ">No</option>" & strLE
				end if
				Response.Write "</select></span></td>" & strLE & _
						"</tr>" & strLE
			end if
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r""><b><span class=""dff ffs"">Signature checkbox<br>checked by default?:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc vam""><span class=""dff dfs"">" & strLE & _
					"<select name=""fSigDefault"">" & strLE
			if strMode = "Register" then
				Response.Write "                      	<option value=""1"" selected>Yes</option>" & strLE & _
						"                      	<option value=""0"">No</option>" & strLE
			else
				Response.Write "                      	<option value=""1""" & chkSelect(rs("M_SIG_DEFAULT"),1) & ">Yes</option>" & strLE & _
						"                      	<option value=""0""" & chkSelect(rs("M_SIG_DEFAULT"),0) & ">No</option>" & strLE
			end if
			Response.Write "</select></span></td>" & strLE & _
					"</tr>" & strLE
		end if
	end if
	if Request.Form("Method_Type") = "Modify" then
		Response.Write "<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">Member Level:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc vat"">" & strLE
		if rs("MEMBER_ID") = intAdminMemberID then
			Response.Write "<span class=""dff dfs"">Administrator</span>" & strLE & _
					"<input type=""hidden"" value=""3"" name=""Level"">" & strLE
		else
			Response.Write "<select value=""1"" name=""Level"">" & strLE & _
					"<option value=""1"""
			if rs("M_LEVEL") = 1 then Response.Write(" selected")
			Response.Write ">Normal User</option>" & strLE & _
					"<option value=""2"""
			if rs("M_LEVEL") = 2 then Response.Write(" selected")
			Response.Write ">Moderator</option>" & strLE & _
					"<option value=""3"""
			if rs("M_LEVEL") = 3 then Response.Write(" selected")
			Response.Write ">Administrator</option>" & strLE & _
					"</select>" & strLE
		end if
		Response.Write "</td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""putc vat nw r""><b><span class=""dff dfs"">User allowed to send<br>e-mail before limit of <br>" & intMaxPostsToEMail & " posts is reached? </span></b></td>" & strLE & _
				"<td class=""putc vab"">" & strLE & _
				"<select name=""allowemail"">" & strLE & _
				"<option value=""1"""
		if rs("M_ALLOWEMAIL") = "1" then Response.Write(" selected")
		Response.Write ">Yes</option>" & strLE & _
				"<option value=""0"""
		if rs("M_ALLOWEMAIL") <> "1" then Response.Write(" selected")
		Response.Write ">No</option>" & strLE & _
				"</select>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE
	end if
	if not(strUseExtendedProfile) then
		Response.Write "<tr>" & strLE & _
				"<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">&nbsp;Contact Info&nbsp;</span></b></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> E-mail Address:&nbsp;</span></b></td>" & strLE & _
				"<td class=""putc""><span class=""dff dfs""><input name=""Email"" size=""25"" maxLength=""50"" value="""
		if strMode <> "Register" then Response.Write(ChkString(rs("M_EMAIL"), "display"))
		Response.Write """>" & strLE & _
				"<input type=""hidden"" name=""Email2"" value="""
		if strMode <> "Register" then Response.Write(rs("M_EMAIL"))
		Response.Write """></span></td>" & strLE & _
				"</tr>" & strLE
		if strMode = "Register" then
			Response.Write "</tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs""><span class=""hlfc"">*</span> E-mail Address Again:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Email3"" size=""25"" maxLength=""50"" value=""""></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		Response.Write "<tr class=""vam"">" & strLE & _
				"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff ffs"">Allow Forum Members<br>to Send you E-Mail?:&nbsp;</span></b></td>" & strLE
		if strMode = "Register" then
			Response.Write "<td class=""putc vam""><span class=""dff dfs"">" & strLE & _
					"<select name=""ReceiveEMail"">" & strLE & _
					"                      	<option value=""1"" selected>Yes</option>" & strLE & _
					"                      	<option value=""0"">No</option>" & strLE & _
					"</select></span></td>" & strLE
		else
			Response.Write "<td class=""putc vam""><span class=""dff ffs"">" & strLE & _
					"<select name=""ReceiveEMail"">" & strLE & _
					"                      	<option value=""1"""
			if rs("M_RECEIVE_EMAIL") <> "0" then Response.Write(" selected")
			Response.Write ">Yes</option>" & strLE & _
					"                      	<option value=""0"""
			if rs("M_RECEIVE_EMAIL") = "0" then Response.Write(" selected")
			Response.Write ">No</option>" & strLE & _
					"</select></span></td>" & strLE
		end if
		Response.Write "</tr>" & strLE
		if strAIM = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">AIM:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""AIM"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_AIM"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strICQ = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">ICQ:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""ICQ"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_ICQ"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strMSN = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">MSN:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""MSN"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_MSN"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strYAHOO = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">YAHOO IM:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""YAHOO"" size=""25"" maxLength=""50"" value="""
			if strMode <> "Register" then Response.Write(ChkString(rs("M_YAHOO"), "display"))
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
	end if
	if (strHomepage + strFavLinks) > 0 and not(strUseExtendedProfile) then
		Response.Write "<tr>" & strLE & _
				"<td class=""ccc c"" colspan=""2""><b><span class=""dff dfs cfc"">Links&nbsp;</span></b></td>" & strLE & _
				"</tr>" & strLE
		if strHomepage = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Homepage:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Homepage"" size=""25"" maxLength=""255"" value="""
			if strMode <> "Register" then
				if rs("M_HOMEPAGE") <> " " and lcase(rs("M_HOMEPAGE")) <> "http://" then Response.Write(ChkString(rs("M_HOMEPAGE"), "display")) else Response.Write("http://")
			else
				Response.Write("http://")
			end if
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
		if strFavLinks = "1" then
			Response.Write "<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">Cool Links:&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Link1"" size=""25"" maxLength=""255"" value="""
			if strMode <> "Register" then
				if rs("M_LINK1") <> " " and lcase(rs("M_LINK1")) <> "http://" then Response.Write(ChkString(rs("M_LINK1"), "display")) else Response.Write("http://")
			else
				Response.Write("http://")
			end if
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td class=""putc vat nw r"" width=""10%""><b><span class=""dff dfs"">&nbsp;</span></b></td>" & strLE & _
					"<td class=""putc""><span class=""dff dfs""><input name=""Link2"" size=""25"" maxLength=""255"" value="""
			if strMode <> "Register" then
				if rs("M_LINK2") <> " " and lcase(rs("M_LINK2")) <> "http://" then Response.Write(ChkString(rs("M_LINK2"), "display")) else Response.Write("http://")
			else
				Response.Write("http://")
			end if
			Response.Write """></span></td>" & strLE & _
					"</tr>" & strLE
		end if
	end if
	Response.Write "</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE

	if strUseExtendedProfile then
		Response.Write "<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</a></span></p>" & strLE & _
				"<table width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw c"" " & strColSpan & ">" & strLE & _
				"<input type=""hidden"" value=""" & cLng(Request.Form("MEMBER_ID")) & """ name=""MEMBER_ID"">" & strLE & _
				"<input type=""submit"" value=""Submit"" name=""Submit1"">" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE
	else
		Response.Write "<table width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw c"" " & strColSpan & ">" & strLE & _
				"<input type=""hidden"" value=""" & cLng(Request.Form("MEMBER_ID")) & """ name=""MEMBER_ID"">" & strLE & _
				"<input type=""submit"" value=""Submit"" name=""Submit1"">" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE
	end if
	on error goto 0
end Sub
%>
