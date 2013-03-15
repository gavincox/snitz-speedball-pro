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
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<%
Response.Write "<table width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs"">" & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Forgot your Password?<br></span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE

if lcase(strEmail) <> "1" then Response.Redirect("default.asp")

if Request.Form("mode") <> "DoIt" and Request.Form("mode") <> "UpdateIt" and trim(Request.QueryString("pwkey")) = "" then
	Call ShowForm
elseif trim(Request.QueryString("pwkey")) <> "" and Request.Form("mode") <> "UpdateIt" then
	key = chkString(Request.QueryString("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>Your password key did not match!</b></span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs hlfc"">Your password key did not match the one that we have in our database.<br>Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br>If this problem persists, please contact the <a href=""mailto:" & strSender & """>Administrator</a> of the forums.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>Your password key did not match!</b></span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs hlfc"">Your password key did not match the one that we have in our database.<br>Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br>If this problem persists, please contact the <a href=""mailto:" & strSender & """>Administrator</a> of the forums.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
	else
		PWMember_ID = rsKey("MEMBER_ID")
		Call showForm2
end if

	rsKey.close
	set rsKey = nothing
elseif trim(Request.Form("pwkey")) <> "" and Request.Form("mode") = "UpdateIt" then
	key = chkString(Request.Form("pwkey"),"SQLString")

	'###Forum_SQL
	strSql = "SELECT M_PWKEY, MEMBER_ID, M_NAME, M_EMAIL "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
	strSql = strSql & " AND M_PWKEY = '" & key & "'"

	set rsKey = my_Conn.Execute (strSql)

	if rsKey.EOF or rsKey.BOF then
		'Error message to user
		Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>Your password key did not match!</b></span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">Your password key did not match the one that we have in our database.<br>Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br>If this problem persists, please contact the <a href=""mailto:" & strSender & """>Administrator</a> of the forums.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
	elseif strComp(key,rsKey("M_PWKEY")) <> 0 then
		'Error message to user
		Response.Write "<p class=""c""><span class=""dff hfs hlfc""><b>Your password key did not match!</b></span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">Your password key did not match the one that we have in our database.<br>Please try submitting your UserName and E-mail Address again by clicking the Forgot your Password? link from the Main page of this forum.<br>If this problem persists, please contact the <a href=""mailto:" & strSender & """>Administrator</a> of the forums.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
	else
		if trim(Request.Form("Password")) = "" then Err_Msg = Err_Msg & "<li>You must choose a Password</li>"
		if Len(Request.Form("Password")) > 25 then Err_Msg = Err_Msg & "<li>Your Password can not be greater than 25 characters</li>"
		if Request.Form("Password") <> Request.Form("Password2") then Err_Msg = Err_Msg & "<li>Your Passwords didn't match.</li>"
		if Err_Msg = "" then
			strEncodedPassword = sha256("" & Request.Form("Password"))
			'Update the user's password
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
			strSql = strSql & " SET M_PASSWORD = '" & chkString(strEncodedPassword,"SQLString") & "'"
			strSql = strSql & ", M_PWKEY = ''"
			strSql = strSql & " WHERE MEMBER_ID = " & cLng(Request.Form("MEMBER_ID"))
			strSql = strSql & " AND M_PWKEY = '" & key & "'"
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		else
			if Err_Msg <> "" then
				Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
					"<table class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
				rsKey.close
				set rsKey = nothing
				Call WriteFooter
				Response.End
			end if
		end if
		Response.Write "<p class=""c""><span class=""dff hfs"">Your Password has been updated!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">You may now login"
		if strAuthType = "db" then Response.Write(" with your UserName and new Password")
		Response.Write ".</span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=default.asp"">" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
	end if
	rsKey.close
	set rsKey = nothing
else
	Err_Msg = ""
	if trim(Request.Form("Name")) = "" then Err_Msg = Err_Msg & "<li>You must enter your UserName</li>"
	if trim(Request.Form("Email")) = "" then Err_Msg = Err_Msg & "<li>You must enter your E-mail Address</li>"
	'## Forum_SQL
	strSql = "SELECT MEMBER_ID, M_NAME, M_EMAIL FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_NAME = '" & ChkString(Trim(Request.Form("Name")), "SQLString") &"'"
	strSql = strSql & " AND M_EMAIL = '" & ChkString(Trim(Request.Form("Email")), "SQLString") &"'"
	set rs = my_Conn.Execute (strSql)
	if rs.BOF and rs.EOF then
		Err_Msg = Err_Msg & "<li>Either the UserName or the E-mail Address you entered does not exist in the database.</li>"
	else
		PWMember_ID    = rs("MEMBER_ID")
		PWMember_Name  = rs("M_NAME")
		PWMember_Email = rs("M_EMAIL")
	end if
	rs.close
	set rs = nothing
	if Err_Msg = "" then
		pwkey = GetKey("none")
		'Update the user Member Level
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS "
		strSql = strSql & " SET M_PWKEY = '" & chkString(pwkey,"SQLString") & "'"
		strSql = strSql & " WHERE MEMBER_ID = " & PWMember_ID
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		if lcase(strEmail) = "1" then
			'## E-mails Message to the Author of this Reply.
			strRecipientsName = PWMember_Name
			strRecipients     = PWMember_Email
			strFrom           = strSender
			strFromName       = strForumTitle
			strsubject        = strForumTitle & " - Forgot Your Password? "
			strMessage        = "Hello " & PWMember_Name & strLE & strLE
			strMessage        = strMessage & "You received this message from " & strForumTitle & " because you have completed the First Step on the ""Forgot Your Password?"" page." & strLE & strLE
			strMessage        = strMessage & "Please click on the link below to proceed to the next step." & strLE & strLE
			strMessage        = strMessage & strForumURL & "password.asp?pwkey=" & pwkey & strLE & strLE
			strMessage        = strMessage & strLE & "If you did not forget your password and received this e-mail in error, then you can just disregard/delete this e-mail, no further action is necessary." & strLE & strLE
%>
			<!--#INCLUDE FILE="inc_mail.asp" -->
<%
		end if
	else
		if Err_Msg <> "" then
			Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
				"<table class=""tc"">" & strLE & _
				"<tr>" & strLE & _
				"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
			Call WriteFooter
			Response.End
		end if
	end if
	Response.Write "<p class=""c""><span class=""dff hfs"">Step One is Complete!</span></p>" & strLE & _
		"<p class=""c""><span class=""dff dfs"">Please follow the instructions in the e-mail that has been sent to <b>" & ChkString(PWMember_Email,"email") & "</b> to complete the next step in this process.</span></p>" & strLE & _
		"<meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & strLE & _
		"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p>" & strLE
end if
Call WriteFooter
Response.End

sub ShowForm()
	Response.Write "<form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input name=""mode"" type=""hidden"" value=""DoIt"">" & strLE & _
		"<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
		"<tr>" & strLE & _
		"<td colspan=""2"" class=""hcc vat c""><b><span class=""dff dfs hfc"">Forgot your Password?</span></b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td colspan=""2"" class=""fcc vat l""><span class=""dff ffs ffc"">This is a 3 step process:" & strLE & _
		"<ul>" & strLE & _
		"<span class=""hlfc""><li><b>First Step:</b><br>Enter your username and e-mail address in the form below to receive an e-mail containing a code to verify that you are who you say you are.</li></span>" & strLE & _
		"<li><b>Second Step:</b><br>Check your e-mail and then click on the link that is provided to return to this page.</li>" & strLE & _
		"<li><b>Third Step:</b><br>Choose your new password.</li>" & strLE & _
		"</ul></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc nw r"" width=""50%""><b><span class=""dff dfs"">&nbsp;UserName:&nbsp;</span></b></td>" & strLE & _
		"<td class=""fcc"" width=""50%""><input type=""text"" name=""Name"" size=""25"" maxLength=""25""></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc nw r"" width=""50%""><b><b><span class=""dff dfs"">E-mail Address:&nbsp;</span></b></td>" & strLE & _
		"<td class=""fcc"" width=""50%""><input type=""text"" name=""Email"" size=""25"" maxLength=""50""></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td colspan=""2"" class=""fcc c""><input type=""submit"" value=""Submit"" id=""Submit1"" name=""Submit1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Reset"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form><br>" & strLE
end sub

sub ShowForm2()
	Response.Write "<form action=""password.asp"" method=""Post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input name=""mode"" type=""hidden"" value=""UpdateIt"">" & strLE & _
		"<input name=""MEMBER_ID"" type=""hidden"" value=""" & PWMember_ID & """>" & strLE & _
		"<input name=""pwkey"" type=""hidden"" value=""" & key & """>" & strLE & _
		"<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""hcc vat c"" colspan=""2""><b><span class=""dff dfs hfc"">Forgot your Password?</span></b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc vat l"" colspan=""2""><span class=""dff ffs ffc"">This is a 3 step process:" & strLE & _
		"<ul>" & strLE & _
		"<li><b>First Step:</b><br>Enter your username and e-mail address in the form below to receive an e-mail containing a code to verify that you are who you say you are. <b>(COMPLETED)</b></li>" & strLE & _
		"<li><b>Second Step:</b><br>Check your e-mail and then click on the link that is provided to return to this page. <b>(COMPLETED)</b></li>" & strLE & _
		"<span class=""hlfc""><li><b>Third Step:</b><br>Choose your new password.</li></span>" & strLE & _
		"</ul></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc nw r"" width=""50%""><b><span class=""dff dfs"">Password:&nbsp;</span></b></td>" & strLE & _
		"<td class=""fcc"" width=""50%""><span class=""dff dfs""><input name=""Password"" type=""Password"" size=""25"" maxLength=""25"" value=""""></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc nw r"" width=""50%""><b><span class=""dff dfs"">Password Again:&nbsp;</span></b></td>" & strLE & _
		"<td class=""fcc"" width=""50%""><span class=""dff dfs""><input name=""Password2"" type=""Password"" maxLength=""25"" size=""25"" value=""""></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc c"" colspan=""2""><input type=""submit"" value=""Submit"" id=""Submit1"" name=""Submit1"">&nbsp;&nbsp;&nbsp;<input type=""reset"" value=""Reset"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form><br>" & strLE
end sub
%>
