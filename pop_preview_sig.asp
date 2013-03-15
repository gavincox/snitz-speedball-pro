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
<%
Response.Write "<script type=""text/javascript"">" & strLE & _
	"function submitPreview()" & strLE & _
	"{" & strLE & _
	"document.previewSig.sig.value = window.opener.document.Form1.Sig.value;" & strLE & _
	"document.previewSig.submit()" & strLE & _
	"}" & strLE & _
	"</script>" & strLE
if request("mode") = "" then
	Response.Write "<form action=""pop_preview_sig.asp"" method=""post"" name=""previewSig"">" & strLE & _
		"<input type=""hidden"" name=""sig"" value="""">" & strLE & _
		"<input type=""hidden"" name=""mode"" value=""display"">" & strLE & _
		"</form>" & strLE & _
		"<script type=""text/javascript"">submitPreview();</script>" & strLE
else
	strSigPreview = trim(request.form("sig"))
	if strSigPreview = "" or IsNull(strSigPreview) then
		if strAllowForumCode = "1" then
			strSigPreview = "[center][b]< There is no text to preview ! >[/b][/center]"
		else
			strSigPreview = "<center><b>< There is no text to preview ! ></b></center>"
		end if
	end if
	Response.Write "<table class=""tc"" width=""100%"" height=""80%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tbc"" width=""100%"" height=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""hcc c"" height=""20""><b><span class=""dff dfs hfc"">Signature Preview</span></b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc vab""><hr noshade class=""ffs""><span class=""dff dfs ffc""><span class=""smt"">" & formatStr(chkString(strSigPreview,"preview")) & "</span></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE
end if
Call WriteFooterShort
Response.End
%>
