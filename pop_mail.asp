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
<!--#INCLUDE FILE="inc_sha256.asp" -->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<%
if strLogonForMail = "1" and mlev = 0 then
	Err_Msg =  "<li>You must be logged on to send a message</li>"
	Response.Write "<table>" & strLE & _
		"<tr>" & strLE & _
		"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE
	Call WriteFooterShort
	Response.End
end if

if Request.QueryString("ID") <> "" and IsNumeric(Request.QueryString("ID")) = True then
	intMemberID = cLng(Request.QueryString("ID"))
else
	intMemberID = 0
end if

if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""

	strSql = "SELECT M_NAME, M_POSTS, M_ALLOWEMAIL FROM " & strMemberTablePrefix & "MEMBERS M"
	strSql = strSql & " WHERE M.MEMBER_ID = " & MemberID

	set rs = my_Conn.Execute (strSql)

	If Not rs.EOF then
		intMPosts = cLng(rs("M_POSTS"))
		intAllowEmail = cInt(rs("M_ALLOWEMAIL"))

		If intMPosts < intMaxPostsToEMail and intAllowEmail <> "1" Then
			Err_Msg        = "<li>" & strNoMaxPostsToEMail & "</li>"
			strSpammerName = RS("M_NAME")
			rs.Close

			strSql = "SELECT M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
			strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

			set rs = my_Conn.Execute (strSql)

			If rs.bof or rs.eof Then
				strDestName = ""
			Else
				strDestname = rs("M_NAME")
			End If
			rs.close

			'Send email to forum admin
			strRecipients = strSender
			strFrom       = strSender
			strFromName   = "Automatic Server Email"
			strSubject    = "Possible Spam Poster"
			strMessage    = "There is a possible spam poster at " & strForumTitle & strLE & strLE
			strMessage    = strMessage & "Member " & strSpammerName & ", with MemberID " & MemberID & ", has been trying to send emails to " & strDestName & ", without having enough posts to be allowed to do it." & strLE & strLE
			strMessage    = strMessage & "He has " & intMPosts & " posts, and should have " & intMaxPostsToEMail & " posts." & strLE & strLE
			strMessage    = strMessage & "Here are the message contents: " & strLE & Request.Form("Msg") & strLE & strLE & strLE & strLE
			strMessage    = strMessage & "This is a message sent automatically by the Spam Control Mod ;)."
			%><!--#INCLUDE FILE="inc_mail.asp" --><%
		End If
	Else
		rs.Close
	End If
End If

'## Forum_SQL
strSql = "SELECT M.M_RECEIVE_EMAIL, M.M_EMAIL, M.M_NAME FROM " & strMemberTablePrefix & "MEMBERS M"
strSql = strSql & " WHERE M.MEMBER_ID = " & intMemberID

set rs = my_Conn.Execute (strSql)

Response.Write "<p><span class=""dff hfs"">Send an E-MAIL Message</span></p>" & strLE

if rs.bof or rs.eof then
	rs.close
	set rs = nothing
	Response.Write "<p><span class=""dff dfs hlfc"">There is no Member with that Member ID</span></p>" & strLE
else
	strRName = ChkString(rs("M_NAME"),"display")
	strREmail = rs("M_EMAIL")
	strRReceiveEmail = rs("M_RECEIVE_EMAIL")

	rs.close
	set rs = nothing

	if mLev > 2 or strRReceiveEmail = "1" then
		if lcase(strEmail) = "1" then
			if Request.QueryString("mode") = "DoIt" then
				if mLev => 2 then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName = rs2("M_NAME")
					YEmail = rs2("M_EMAIL")
					set rs2 = nothing
				else
					YName = Request.Form("YName")
					YEmail = Request.Form("YEmail")
					if YName = "" then
						Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
					end if
					if YEmail = "" then
						Err_Msg = Err_Msg & "<li>You must give your e-mail address</li>"
					else
						if EmailField(YEmail) = 0 then
							Err_Msg = Err_Msg & "<li>You must enter a valid e-mail address</li>"
						end if
					end if
				end if
				if Request.Form("Msg") = "" then
					Err_Msg = Err_Msg & "<li>You must enter a message</li>"
				end if
				'##  E-mails Message to the Author of this Reply.
				if (Err_Msg = "") then
					strRecipientsName = strRName
					strRecipients = strREmail
					strFrom = YEmail
					strFromName = YName
					strSubject = "Sent From " & strForumTitle & " by " & YName
					strMessage = "Hello " & strRName & strLE & strLE
					strMessage = strMessage & "You received the following message from: " & YName & " (" & YEmail & ") " & strLE & strLE
					strMessage = strMessage & "At: " & strForumURL & strLE & strLE
					strMessage = strMessage & Request.Form("Msg") & strLE & strLE

					if strFrom <> "" then
						strSender = strFrom
					end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
					Response.Write "<p><span class=""dff hfs"">E-mail has been sent</span></p>" & strLE
				else
					Response.Write "<p><span class=""dff hfs hlfc"">There Was A Problem With Your E-mail</span></p>" & strLE
					Response.Write "<table>" & strLE & _
						"<tr>" & strLE & _
						"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
						"</tr>" & strLE & _
						"</table>" & strLE & _
						"<p><span class=""dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
					Call WriteFooterShort
					Response.End
				end if
			else
				Err_Msg = ""
				if trim(strREmail) <> "" then
					strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
					strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
					strSql = strSql & " WHERE MEMBER_ID = " & MemberID

					set rs2 = my_conn.Execute (strSql)
					YName  = ""
					YEmail = ""

					if (rs2.EOF or rs2.BOF)  then
						if strLogonForMail <> "0" then
							Err_Msg = Err_Msg & "<li>You must be logged on to send a message</li>"

							Response.Write "<table>" & strLE & _
								"<tr>" & strLE & _
								"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
								"</tr>" & strLE & _
								"</table>" & strLE
							Call WriteFooterShort
							Response.End
						end if
					else
						YName  = Trim("" & rs2("M_NAME"))
						YEmail = Trim("" & rs2("M_EMAIL"))
					end if
					rs2.close
					set rs2 = nothing

					Response.Write "<form action=""pop_mail.asp?mode=DoIt&id=" & intMemberID & """ method=""Post"" id=""Form1"" name=""Form1"">" & strLE & _
						"<table width=""90%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
						"<tr>" & strLE & _
						"<td class=""pubc"">" & strLE & _
						"<table width=""100%"" cellspacing=""1"" cellpadding=""1"">" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc nw r""><b><span class=""dff dfs"">Send To Name:</span></b></td>" & strLE & _
						"<td class=""putc""><span class=""dff dfs"">" & strRName & "</span></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc nw r""><b><span class=""dff dfs"">Your Name:</span></b></td>" & strLE & _
						"<td class=""putc"">"
					if YName = "" then
						Response.Write "<input name=""YName"" type=""text"" value=""" & YName & """ size=""25"">"
					else
						Response.Write "<span class=""dff dfs"">" & YName & "</span>" & strLE
					end if
					Response.Write "</td></tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc nw r""><b><span class=""dff dfs"">Your E-mail:</span></b></td>" & strLE & _
						"<td class=""putc"">"
					if YEmail = "" then
						Response.Write "<input name=""YEmail"" type=""text"" value=""" & YEmail & """ size=""25"">"
					else
						Response.Write "<span class=""dff dfs"">" & YEmail & "</span>"
					end if
					Response.Write "</td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc"" colspan=""2""><b><span class=""dff dfs"">Message:</span></b></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc"" colspan=""2""><textarea name=""Msg"" cols=""40"" rows=""5""></textarea></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""putc c"" colspan=""2""><input type=""Submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
						"</tr>" & strLE & _
						"</table>" & strLE & _
						"</td>" & strLE & _
						"</tr>" & strLE & _
						"</table>" & strLE & _
						"</form>" & strLE
				else
					Response.Write "<p><span class=""dff dfs hlfc"">No E-mail address is available for this user.</span></p>" & strLE
				end if
			end if
		else
			Response.Write "<p><span class=""dff dfs"">Click to send <a href=""mailto:" & chkString(strREmail,"display") & """>" & strRName & "</a> an e-mail</span></p>" & strLE
		end if
	else
		Response.Write "<p><span class=""dff dfs hlfc"">This Member does not wish to receive e-mail.</span></p>" & strLE
	end if
end if
Call WriteFooterShort
Response.End
%>
