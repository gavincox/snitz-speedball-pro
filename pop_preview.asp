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
<!--#INCLUDE FILE="inc_header_short.asp"-->
<!--#INCLUDE FILE="inc_func_member.asp" -->
<!--#INCLUDE FILE="inc_func_posting.asp"-->
<!--#INCLUDE FILE="inc_func_secure.asp"-->
<%
Response.Write "<script type=""text/javascript"">" & strLE & _
	"function submitPreview()" & strLE & _
	"{" & strLE & _
	"if (window.opener.document.PostTopic.Subject) {" & strLE & _
	"document.previewTopic.subject.value = window.opener.document.PostTopic.Subject.value;" & strLE & _
	"}" & strLE & _
	"document.previewTopic.message.value = window.opener.document.PostTopic.Message.value;" & strLE & _
	"if (window.opener.document.PostTopic.Sig) {" & strLE & _
	"if (window.opener.document.PostTopic.Sig.checked) {" & strLE & _
	"document.previewTopic.sig.value = ""yes""" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"if (window.opener.document.PostTopic.Author) {" & strLE & _
	"document.previewTopic.author.value = window.opener.document.PostTopic.Author.value;" & strLE & _
	"}" & strLE & _
	"document.previewTopic.submit()" & strLE & _
	"}" & strLE & _
	"</script>" & strLE
if request("mode") = "" then
	Response.Write "<form action=""pop_preview.asp"" method=""post"" name=""previewTopic"">" & strLE & _
		"<input type=""hidden"" name=""subject"" value="""">" & strLE & _
		"<input type=""hidden"" name=""message"" value="""">" & strLE & _
		"<input type=""hidden"" name=""sig"" value="""">" & strLE & _
		"<input type=""hidden"" name=""author"" value="""">" & strLE & _
		"<input type=""hidden"" name=""mode"" value=""display"">" & strLE & _
		"</form>" & strLE & _
		"<script type=""text/javascript"">submitPreview();</script>" & strLE
else
	CColor = strForumCellColor
	strSubjectPreview = trim(Request.Form("subject"))
	strMessagePreview = trim(Request.Form("message"))
	if strMessagePreview = "" or IsNull(strMessagePreview) then
		if strAllowForumCode = "1" then
			strMessagePreview = "[center][b]< There is no text to preview ! >[/b][/center]"
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		else
			strMessagePreview = "<center><b>< There is no text to preview ! ></b></center>"
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		end if
	else
		if Request.Form("author") = "" or isNull(Request.Form("author")) then
			strSigAuthor = strDBNTUserName
		else
			strSigAuthor = ChkString(getMemberName(Request.Form("author")),"SQLString")
		end if
		if Request.Form("sig") = "yes" and trim(GetSig(strSigAuthor)) <> "" then
			if strDSignatures = "1" then
				strMessagePreview = formatStr(chkString(strMessagePreview,"preview")) & "<hr noshade class=""ffs"">" & formatStr(chkString(cleancode(GetSig(strSigAuthor)),"preview"))
			else
				strMessagePreview = strMessagePreview & strLE & strLE & CleanCode(GetSig(strSigAuthor))
				strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
			end if
		else
			strMessagePreview = formatStr(chkString(strMessagePreview,"preview"))
		end if
	end if
	if strSubjectPreview = "" or IsNull(strSubjectPreview) then
		strPreviewTitle = "Message Preview"
	else
		strPreviewTitle = "Message Preview - " & chkString(strSubjectPreview,"display")
	end if

	Response.Write "<table class=""tc"" width=""100%"" height=""80%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tbc"" width=""100%"" height=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""hcc c"" height=""20""><b><span class=""dff dfs hfc"">" & strPreviewTitle & "</span></b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc vat""><span class=""dff dfs ffc""><span class=""smt"">" & strMessagePreview & "</span></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE
end if
Call WriteFooterShort
Response.End
%>
