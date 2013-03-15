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
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Email&nbsp;Server&nbsp;Configuration" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE

if Request.Form("Method_Type") = "Write_Configuration" then
	Err_Msg = ""
	if Request.Form("strMailServer") = "" and Request.Form("strMailMode") <> "cdonts" and Request.Form("strEmail") = "1" then
		Err_Msg = Err_Msg & "<li>You Must Enter the Address of your Mail Server</li>"
	end if
	if ((lcase(left(Request.Form("strMailServer"), 7)) = "http://") or (lcase(left(Request.Form("strMailServer"), 8)) = "https://")) and Request.Form("strEmail") = "1" then
		Err_Msg = Err_Msg & "<li>Do not prefix the Mail Server Address with <b>http://</b>, <b>https://</b> or <b>file://</b></li>"
	end if
	if Request.Form("strSender") = "" then
		Err_Msg = Err_Msg & "<li>You Must Enter the E-mail Address of the Forum Administrator</li>"
	else
		if EmailField(Request.Form("strSender")) = 0 and Request.Form("strSender") <> "" then
			Err_Msg = Err_Msg & "<li>You Must enter a valid E-mail Address for the Forum Administrator</li>"
		end if
	end if
	if Request.Form("strRestrictReg") = 1 and Request.Form("strEmailVal") = 0 then
		Err_Msg = Err_Msg & "<li>Email Validation must be enabled in order to enable the Restrict Registration Option</li>"
	end if
	if not IsNumeric(Request.Form("intMaxPostsToEMail")) then
		Err_Msg = Err_Msg & "<li>Number of posts to allow sending e-mail must be a number</li>"
	end if

	if Err_Msg = "" then
		'## Forum_SQL
		for each key in Request.Form
			if left(key,3) = "str" or left(key,3) = "int" then
				strDummy = SetConfigValue(1, key, ChkString(Request.Form(key),"SQLString"))
			end if
		next
		Application(strCookieURL & "ConfigLoaded") = ""

		Response.Write "<p class=""c""><span class=""dff hfs"">Configuration Posted!</span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_home.asp"">" & strLE & _
			"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
			"<p class=""c""><a href=""admin_home.asp"">Back To Admin Home</span></a></p>" & strLE
	else
		Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
			"<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
	end if
else
	Dim theComponent(20)
	Dim theComponentName(20)
	Dim theComponentValue(20)

	'## the components
	theComponent(0) = "ABMailer.Mailman"
	theComponent(1) = "Persits.MailSender"
	theComponent(2) = "SMTPsvg.Mailer"
	theComponent(3) = "SMTPsvg.Mailer"
	theComponent(4) = "CDONTS.NewMail"
	theComponent(5) = "CDONTS.NewMail"
	theComponent(6) = "CDO.Message"
	theComponent(7) = "dkQmail.Qmail"
	theComponent(8) = "Dundas.Mailer"
	theComponent(9) = "Dundas.Mailer"
	theComponent(10) = "Innoveda.MailSender"
	theComponent(11) = "Geocel.Mailer"
	theComponent(12) = "iismail.iismail.1"
	theComponent(13) = "Jmail.smtpmail"
	theComponent(14) = "Jmail.Message"
	theComponent(15) = "MDUserCom.MDUser"
	theComponent(16) = "ASPMail.ASPMailCtrl.1"
	theComponent(17) = "ocxQmail.ocxQmailCtrl.1"
	theComponent(18) = "SoftArtisans.SMTPMail"
	theComponent(19) = "SmtpMail.SmtpMail.1"
	theComponent(20) = "VSEmail.SMTPSendMail"

	'## the name of the components
	theComponentName(0) = "ABMailer v2.2+"
	theComponentName(1) = "ASPEMail"
	theComponentName(2) = "ASPMail"
	theComponentName(3) = "ASPQMail"
	theComponentName(4) = "CDONTS (IIS 3/4/5)"
	theComponentName(5) = "Chili!Mail (Chili!Soft ASP)"
	theComponentName(6) = "CDOSYS (IIS 5/5.1/6)"
	theComponentName(7) = "dkQMail"
	theComponentName(8) = "Dundas Mail (QuickSend)"
	theComponentName(9) = "Dundas Mail (SendMail)"
	theComponentName(10) = "FreeMailSender"
	theComponentName(11) = "GeoCel"
	theComponentName(12) = "IISMail"
	theComponentName(13) = "JMail 3.x"
	theComponentName(14) = "JMail 4.x"
	theComponentName(15) = "MDaemon"
	theComponentName(16) = "OCXMail"
	theComponentName(17) = "OCXQMail"
	theComponentName(18) = "SA-Smtp Mail"
	theComponentName(19) = "SMTP"
	theComponentName(20) = "VSEmail"

	'## the value of the components
	theComponentValue(0) = "abmailer"
	theComponentValue(1) = "aspemail"
	theComponentValue(2) = "aspmail"
	theComponentValue(3) = "aspqmail"
	theComponentValue(4) = "cdonts"
	theComponentValue(5) = "chilicdonts"
	theComponentValue(6) = "cdosys"
	theComponentValue(7) = "dkqmail"
	theComponentValue(8) = "dundasmailq"
	theComponentValue(9) = "dundasmails"
	theComponentValue(10) = "freemailsender"
	theComponentValue(11) = "geocel"
	theComponentValue(12) = "iismail"
	theComponentValue(13) = "jmail"
	theComponentValue(14) = "jmail4"
	theComponentValue(15) = "mdaemon"
	theComponentValue(16) = "ocxmail"
	theComponentValue(17) = "ocxqmail"
	theComponentValue(18) = "sasmtpmail"
	theComponentValue(19) = "smtp"
	theComponentValue(20) = "vsemail"

	Response.Write "<form action=""admin_config_email.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
		"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
		"<table class=""admin"">" & strLE & _
		"<tr>" & strLE & _
		"<th colspan=""2""><b>E-mail Server Configuration</b></th>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Select E-mail Component</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<select name=""strMailMode"">" & strLE
	dim i, j
	j = 0
	for i=0 to UBound(theComponent)
		if IsObjInstalled(theComponent(i)) then
			Response.Write "<option value=""" & theComponentValue(i) & """" & chkSelect(strMailMode,theComponentValue(i)) & ">" & theComponentName(i) & "</option>" & strLE
		else
			j = j + 1
		end if
	next
	if j > UBound(theComponent) then
		Response.Write "<option value=""None"">No Compatible Component Found</option>" & strLE
	end if

	Response.Write "</select>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>E-mail Mode</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strEmail"" value=""1"""
	if j > UBound(theComponent) then Response.Write(" disabled") else if lcase(strEmail) <> "0" then Response.Write(" checked")
	Response.Write "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strEmail"" value=""0"""
	if j > UBound(theComponent) then Response.Write(" checked") else if lcase(strEmail) = "0" then Response.Write(" checked")
	Response.Write "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#email')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>E-mail Server Address</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""text"" name=""strMailServer"" size=""40"" value=""" & strMailServer & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#mailserver')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Administrator E-mail Address</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""text"" name=""strSender"" size=""40"" value=""" & strSender & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#sender')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Require Unique E-mail</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""1""" & chkRadio(strUniqueEmail,1,true) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strUniqueEmail"" value=""0""" & chkRadio(strUniqueEmail,1,false) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#UniqueEmail')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>E-mail Validation</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strEmailVal"" value=""1""" & chkRadio(strEmailVal,1,true) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strEmailVal"" value=""0""" & chkRadio(strEmailVal,1,false) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#EmailVal')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Filter known spam domains</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFilterEMailAddresses"" value=""1""" & chkRadio(strFilterEMailAddresses,1,true) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strFilterEMailAddresses"" value=""0""" & chkRadio(strFilterEMailAddresses,1,false) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#EmailFilter')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Restrict Registration</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""1""" & chkRadio(strRestrictReg,1,true) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strRestrictReg"" value=""0""" & chkRadio(strRestrictReg,1,false) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#RestrictReg')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Require Logon for sending Mail</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""1""" & chkRadio(strLogonForMail,1,true) & "> On&nbsp;" & strLE & _
		"<input type=""radio"" class=""radio"" name=""strLogonForMail"" value=""0""" & chkRadio(strLogonForMail,1,false) & "> Off" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#LogonForMail')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Number of posts to allow sending e-mail</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""text"" name=""intMaxPostsToEMail"" size=""5"" maxlength=""10"" value=""" & intMaxPostsToEMail & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#MaxPostsToEMail')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""r""><b>Error if they don't have enough posts</b>&nbsp;</td>" & strLE & _
		"<td>" & strLE & _
		"<input type=""text"" name=""strNoMaxPostsToEMail"" size=""40"" maxlength=""255"" value=""" & strNoMaxPostsToEMail & """>" & strLE & _
		"<a href=""JavaScript:openWindow3('pop_config_help.asp?mode=email#NoMaxPostsToEMail')"">" & getCurrentIcon(strIconSmileQuestion,"","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""c"" colspan=""2""><input type=""submit"" value=""Submit New Config"" name=""submit1""> <input type=""reset"" value=""Reset Old Values"" id=""reset1"" name=""reset1""></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</form>" & strLE
end if
WriteFooter
Response.End

function IsObjInstalled(strClassString)
	on error resume next
	'## initialize default values
	IsObjInstalled = false
	Err.Clear
	'## testing code
	dim xTestObj
	set xTestObj = Server.CreateObject(strClassString)
	if Err.Number = 0 then
		IsObjInstalled = true
	end if
	'## cleanup
	set xTestObj = nothing
	Err.Clear
	on error goto 0
end function
%>
