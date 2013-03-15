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
if Request.QueryString("mode") = "DoIt" then
	Err_Msg = ""

	if strLogonForMail <> "0" and (MemberID < 1 or isNull(MemberID)) then
		Err_Msg = Err_Msg & "<li>You Must be logged on to send a message</li>"
	end if
	if (Request.Form("YName") = "") then
		Err_Msg = Err_Msg & "<li>You must enter your name!</li>"
	end if
	if (Request.Form("YEmail") = "") then
		Err_Msg = Err_Msg & "<li>You Must give your e-mail address</li>"
	else
		if (EmailField(Request.Form("YEmail")) = 0) then
			Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address</li>"
		end if
	end if
	if (Request.Form("Name") = "") then
		Err_Msg = Err_Msg & "<li>You must enter the recipients name</li>"
	end if
	if (Request.Form("Email") = "") then
		Err_Msg = Err_Msg & "<li>You Must enter the recipients e-mail address</li>"
	else
		if (EmailField(Request.Form("Email")) = 0) then
			Err_Msg = Err_Msg & "<li>You Must enter a valid e-mail address for the recipient</li>"
		end if
	end if
	if (Request.Form("Msg") = "") then
		Err_Msg = Err_Msg & "<li>You Must enter a message</li>"
	end if
	if lcase(strEmail) = "1" then
		if (Err_Msg = "") then
			strRecipientsName = Request.Form("Name")
			strRecipients = Request.Form("Email")
			strSubject = "From: " & Request.Form("YName") & " Interesting Page"
			strMessage = "Hello " & Request.Form("Name") & strLE & strLE
			strMessage = strMessage & Request.Form("Msg") & strLE & strLE
			strMessage = strMessage & "You received this from : " & Request.Form("YName") & " (" & Request.Form("YEmail") & ") "

			if Request.Form("YEmail") <> "" then
				strSender = Request.Form("YEmail")
			end if
%>
<!--#INCLUDE FILE="inc_mail.asp" -->
<%
			Response.Write("<p><span class=""dff hfs"">E-mail has been sent</span></p>" & strLE)
		else
			Response.Write("<p><span class=""dff hfs hlfc"">There Was A Problem With Your E-mail</span></p>" & strLE)
			Response.Write "<table>" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE
			Response.Write("<p><span class=""dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE)
		end if
	end if
else
	'## Forum_SQL
	strSql =  "SELECT M_NAME, M_USERNAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS"
	strSql = strSql & " WHERE " & strDBNTSQLName & " = '" & chkString(strDBNTUserName,"SQLString") & "'"

	set rs = my_conn.Execute (strSql)
	YName = ""
	YEmail = ""

	if (rs.EOF or rs.BOF)  then
		if strLogonForMail <> "0" then
			Err_Msg = Err_Msg & "<li>You Must be logged on to send a message</li>"
			Response.Write("<p><span class=""dff hfs hlfc"">There Was A Problem With Your E-mail</span></p>" & strLE)
			Response.Write "<table>" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE
			set rs = nothing
			Response.Write("<p><span class=""dfs""><a href=""JavaScript:onClick= window.close()"">Close Window</A></span></p>" & strLE)
			Response.End
		end if
	else
	  	YName = Trim("" & rs("M_NAME"))
	  	YEmail = Trim("" & rs("M_EMAIL"))
	end if

	rs.close
	set rs = nothing

	Response.Write("<p><span class=""dff hfs"">Send Topic to a Friend</span></p>" & strLE)

	Response.Write "<form action=""pop_send_to_friend.asp?mode=DoIt"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
			"<input type=""hidden"" name=""Page"" value=""" & Request.QueryString & """>" & strLE & _
			"<table width=""90%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""pubc"">" & strLE & _
			"<table width=""100%"" cellspacing=""1"" cellpadding=""1"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">Send To Name:</span></b></td>" & strLE & _
			"<td class=""putc""><input type=""text"" name=""Name"" size=""25""></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">Send To E-mail:</span></b></td>" & strLE & _
			"<td class=""putc""><input type=""text"" name=""Email"" size=""25""></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">Your Name:</span></b></td>" & strLE & _
			"<td class=""putc""><input name=""YName"" type="""
			if YName <> "" then
				Response.Write("hidden")
			else
				Response.Write("text")
			end if
	Response.Write """ value=""" & YName & """ size=""25""><span class=""dff dfs"">"
			if YName <> "" then
				Response.Write(YName)
			end if
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw r""><b><span class=""dff dfs"">Your E-mail:</span></b></td>" & strLE & _
			"<td class=""putc""><input name=""YEmail"" type="""
			if YEmail <> "" then
				Response.Write("hidden")
			else
				Response.Write("text")
			end if
	Response.Write """ value=""" & YEmail & """ size=""25""><span class=""dff dfs"">"
			if YEmail <> "" then
				Response.Write(YEmail)
			end if
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc nw"" colspan=""2""><b><span class=""dff dfs"">Message:</span></b></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc c"" colspan=""2""><textarea name=""Msg"" cols=""40"" rows=""5"" readonly>I thought you might be interested in this post:" & strLE & strLE & Request.QueryString("url") & "</textarea></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""putc c"" colspan=""2""><input type=""submit"" value=""Send"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE
end if
WriteFooterShort
Response.End
%>
