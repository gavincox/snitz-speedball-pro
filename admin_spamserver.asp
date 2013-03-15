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
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="inc_func_admin.asp" -->
<%
If Session(strCookieURL & "Approval") <> "15916941253" Then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
End If
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs w50"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Blocked E-Mail Domains<br></span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & strLE
Dim strMethodType
If Request.Form("Method_Type") <> "" Then
	strMethodType = LCase(Trim(Request.Form("Method_Type")))
Else
	If Request.QueryString("Method_Type") <> "" Then
		strMethodType = LCase(Trim(Request.QueryString("Method_Type")))
	Else
		strMethodType = "blank"
	End If
End If
Select Case strMethodType
	Case "add"
		'### Add the e-mail domain to the list
		Dim strSpammServer : strSpammServer = LCase(Trim(chkString(Request.Form("SpamServer"),"sqlstring")))
		Err_Msg = ""
		if strSpammServer = "" then Err_Msg = Err_Msg & "<li>You need to enter an address to block.</li>"
		if (Instr(strSpammServer, " ") > 0 ) then Err_Msg = Err_Msg & "<li>You cannot have spaces in your address.</li>"
		'Comment out down to the next comment to let it take me@example.com and/or .ex as well
		If Left(strSpammServer,1) = "@" Then
			If InStr(1,strSpammServer,".",vbTextCompare) = 0 Then Err_Msg = Err_Msg & "<li>You need to have a TLD (.com, .net, .whatever) in your address.</li>"
		Else
			If InStr(1,strSpammServer,"@",vbTextCompare) <> 0 Then
				Err_Msg = Err_Msg & "<li>You can only enter a domain (@example.com), not a specific address (me@example.com).</li>"
			Else
				If InStr(1,strSpammServer,".",vbTextCompare) = 0 Then
					Err_Msg = Err_Msg & "<li>You need to enter a valid domain (@example.com)</li>"
				Else
					strSpammServer = "@" & strSpammServer
				End If
			End If
		End if
		'Comment out up to the previous comment to let it take me@example.com and/or .ex as well
		If Err_Msg = "" Then
			strSQL = "SELECT SPAM_ID, SPAM_SERVER FROM " & strFilterTablePrefix & "SPAM_MAIL WHERE SPAM_SERVER = '"& strSpammServer &"'"
			set rs = my_conn.execute(strSQL)
			If Not rs.EOF And Not rs.BOF Then Err_Msg = Err_Msg & "<li>'" & strSpammServer & "' is already in the list</li>"
			rs.Close
			Set rs = Nothing
		End If
		if Err_Msg = "" then
			strSQL = "INSERT INTO " & strFilterTablePrefix & "SPAM_MAIL (SPAM_SERVER) VALUES ('" & strSpammServer & "')"
			my_Conn.Execute (strSql)
			Response.Write "<p class=""c""><span class=""dff dfs"">" & _
				"E-Mail domain added<br><br><a href=""admin_spamserver.asp"">Back to the list</a></span></p>" & strLE & _
				"<meta http-equiv=""Refresh"" content=""8; URL=admin_spamserver.asp"">" & strLE
		else
			Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
				"<table class=""tc"">" & strLE & _
				"<tr>" & strLE & _
				"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
		end if
		'############### Add SpamServer end
	Case "delete"
		'### Delete the domain from the list if its a valid ID
		If Request.QueryString("id") <> "" And IsNumeric(Request.QueryString("id")) Then
			Dim intSpamID : intSpamID = cLng(Request.QueryString("id"))
			strSQL = "DELETE FROM " & strFilterTablePrefix & "SPAM_MAIL WHERE SPAM_ID = " & intSpamID
			my_Conn.Execute (strSql)
			Response.Write "<p class=""c""><span class=""dff dfs"">" & _
				"E-Mail domain deleted<br><br><a href=""admin_spamserver.asp"">Back to the list</a></span></p>" & strLE & _
				"<meta http-equiv=""Refresh"" content=""8; URL=admin_spamserver.asp"">" & strLE
		Else
			'Not a numerid ID, forward them to the admin page
			Response.Redirect("admin_spamserver.asp")
		End If
		'############### Delete SpamServer end
	Case Else
		'### Show the admin form and the list of currently blocked domains
		Response.Write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th colspan=""2""><b>Blocked E-Mail Domains</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr class=""vat"">" & strLE & _
			"<td class=""r""><span class=""dff dfs"">Block Email Domain&nbsp;</span><br>" & _
			"<span class=""ffs"">(like '@example.com')&nbsp;</td>" & strLE & _
			"<td>" & _
			"<form action=""admin_spamserver.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
			"<input type=""hidden"" name=""Method_Type"" value=""add"">" & strLE & _
			"<input type=""text"" name=""SpamServer"" size=""50"" maxlength=""255"" value="""">&nbsp;<input type=""submit"" value=""Block"" id=""submit1"" name=""submit1"">" & strLE & _
			"</form></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr class=""vat"">" & strLE & _
			"<td class=""r"">Email Domains blocked&nbsp;</td>"& strLE & _
			"<td>" & strLE
		Set rs = my_conn.Execute("SELECT SPAM_ID, SPAM_SERVER FROM " & strFilterTablePrefix & "SPAM_MAIL")
		If rs.EOF or rs.BOF Then
			Response.Write "No blocked e-mail domains found" & strLE
		Else
			Do Until rs.EOF
				Response.Write rs("SPAM_SERVER") & "&nbsp;<a href=""#"" onClick=""JavaScript:if(window.confirm('Delete Spam Server on this domain?')){window.location=('admin_spamserver.asp?Method_Type=delete&id=" & rs("SPAM_ID") & "');}"">" & getCurrentIcon(strIconTrashcan,"Remove block","class=""vam""") & "</a><br>" & strLE
				rs.MoveNext
			Loop
		End If
		rs.Close
		Set rs = Nothing
		Response.Write "</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE
End Select
Call WriteFooter
Response.end
%>