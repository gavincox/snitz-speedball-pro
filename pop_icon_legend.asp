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
<%
strSmileCode = array("[:)]","[:D]","[8D]","[:I]","[:p]","[}:)]","[;)]","[:o)]","[B)]","[8]","[:(]","[8)]","[:0]","[:(!]","[xx(]","[|)]","[:X]","[^]","[V]","[?]")
strSmileDesc = array("smile","big smile","cool","blush","tongue","evil","wink","clown","black eye","eightball","frown","shy","shocked","angry","dead","sleepy","kisses","approve","disapprove","question")
strSmileName = array(strIconSmile,strIconSmileBig,strIconSmileCool,strIconSmileBlush,strIconSmileTongue,strIconSmileEvil,strIconSmileWink,strIconSmileClown,strIconSmileBlackeye,strIconSmile8ball,strIconSmileSad,strIconSmileShy,strIconSmileShock,strIconSmileAngry,strIconSmileDead,strIconSmileSleepy,strIconSmileKisses,strIconSmileApprove,strIconSmileDisapprove,strIconSmileQuestion)

Response.Write "<script type=""text/javascript"">" & strLE & _
	"<!-- " & vbNewLine & _
    "function insertsmilie(smilieface) {" & strLE & _
	"if (window.opener.document.PostTopic.Message.createTextRange && window.opener.document.PostTopic.Message.caretPos) {" & strLE & _
	"var caretPos = window.opener.document.PostTopic.Message.caretPos;" & strLE & _
	"caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ? smilieface + ' ' : smilieface;" & strLE & _
	"window.opener.document.PostTopic.Message.focus();" & strLE & _
	"} else {" & strLE & _
	"window.opener.document.PostTopic.Message.value+=smilieface;" & strLE & _
	"window.opener.document.PostTopic.Message.focus();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"// -->" & strLE & _
	"</script>" & strLE & _
	"<table class=""tc"" width=""95%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a name=""smilies""></a><span class=""dff dfs cfc""><b>Smilies</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""fcc"">" & strLE & _
	"<p><span class=""dff dfs ffc"">" & strLE & _
	"You've probably seen others use smilies before in e-mail messages or other bulletin " & strLE & _
	"board posts. Smilies are keyboard characters used to convey an emotion, such as a smile " & strLE & _
	getCurrentIcon(strIconSmile,"","class=""vam""") & " or a frown " & strLE & _
	getCurrentIcon(strIconSmileSad,"","class=""vam""") & ". This bulletin board " & strLE & _
	"automatically converts certain text to a graphical representation when it is " & strLE & _
	"inserted between brackets [].&nbsp; Here are the smilies that are currently " & strLE & _
	"supported by " & strForumTitle & ":<br>" & strLE & _
	"<table class=""tc"" cellpadding=""5"">" & strLE & _
	"<tr class=""vat"">" & strLE & _
	"<td>" & strLE & _
	"<table class=""tc"">" & strLE
for sm = 0 to 9
	Response.Write "<tr>" & strLE & _
		"<td class=""fcc""><a href=""Javascript:insertsmilie('" & strSmileCode(sm) & "');"">" & getCurrentIcon(strSmileName(sm),"","class=""vam""") & "</a></td>" & strLE & _
		"<td class=""fcc""><span class=""dff dfs"">" & strSmileDesc(sm) & "</span></td>" & strLE & _
		"<td class=""fcc""><span class=""dff dfs"">" & strSmileCode(sm) & "</span></td>" & strLE & _
		"</tr>" & strLE
next
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tc"">>" & strLE
for sm = 10 to 19
	Response.Write "<tr>" & strLE & _
		"<td class=""fcc""><a href=""Javascript:insertsmilie('" & strSmileCode(sm) & "');"">" & getCurrentIcon(strSmileName(sm),"","class=""vam""") & "</a></td>" & strLE & _
		"<td class=""fcc""><span class=""dff dfs"">" & strSmileDesc(sm) & "</span></td>" & strLE & _
		"<td class=""fcc""><span class=""dff dfs"">" & strSmileCode(sm) & "</span></td>" & strLE & _
		"</tr>" & strLE
next
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table></p>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
Call WriteFooterShort
Response.End
%>
