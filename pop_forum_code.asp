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
Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a name=""format""></a><span class=""dff dfs cfc""><b>How to format text with Bold, Italic, Quote, etc...</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""fcc"">" & strLE & _
	"<p><span class=""dff dfs ffc"">" & strLE & _
	"There are several Forum Codes you may use to change the appearance " & strLE & _
	"of your text.&nbsp; Following is the list of codes currently available:</p>" & strLE & _
	"<blockquote>" & strLE & _
	"<p><b>Bold:</b> Enclose your text with [b] and [/b] .&nbsp; <i>Example:</i> This is <b>[b]</b>bold<b>[/b]</b> text. = This is <b>bold</b> text.</p>" & strLE & _
	"<p><i>Italic:</i> Enclose your text with [i] and [/i] .&nbsp; <i>Example:</i> This is <b>[i]</b>italic<b>[/i]</b> text. = This is <i>italic</i> text.</p>" & strLE & _
	"<p><u>Underline:</u> Enclose your text with [u] and [/u]. <i>Example:</i> This is <b>[u]</b>underline<b>[/u]</b> text. =  This is <u>underline</u> text.</p>" & strLE & _
	"<p><b>Aligning Text Left:</b><br>" & strLE & _
	"Enclose your text with [left] and [/left]" & strLE & _
	"</p>" & strLE & _
	"<p><b>Aligning Text Center:</b><br>" & strLE & _
	"Enclose your text with [center] and [/center]" & strLE & _
	"</p>" & strLE & _
	"<p><b>Aligning Text Right:</b><br>" & strLE & _
	"Enclose your text with [right] and [/right]" & strLE & _
	"</p>" & strLE & _
	"<p><b>Striking Text:</b><br>" & strLE & _
	"Enclose your text with [s] and [/s]<br>" & strLE & _
	"<i>Example:</i> <b>[s]</b>mistake<b>[/s]</b> = <s>mistake</s>" & strLE & _
	"</p>" & strLE & _
	"<p><b>Horizontal Rule:</b><br>" & strLE & _
	"Place a horizontal line in your post with [hr]<br>" & strLE & _
	"<i>Example:</i> <b>[hr]</b> = <hr noshade size=""1"">" & strLE & _
	"</p>" & strLE & _
	"<p>&nbsp; </p>" & strLE & _
	"<p><b>Font Colors:</b><br>" & strLE & _
	"Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>] <br>" & strLE & _
	"<i>Example:</i> <b>[red]</b>Text<b>[/red]</b> = <span color=""red"">Text</font id=""red""><br>" & strLE & _
	"<i>Example:</i> <b>[blue]</b>Text<b>[/blue]</b> = <span color=""blue"">Text</font id=""blue""><br>" & strLE & _
	"<i>Example:</i> <b>[pink]</b>Text<b>[/pink]</b> = <span color=""pink"">Text</font id=""pink""><br>" & strLE & _
	"<i>Example:</i> <b>[brown]</b>Text<b>[/brown]</b> = <span color=""brown"">Text</font id=""brown""><br>" & strLE & _
	"<i>Example:</i> <b>[black]</b>Text<b>[/black]</b> = <span color=""black"">Text</font id=""black""><br>" & strLE & _
	"<i>Example:</i> <b>[orange]</b>Text<b>[/orange]</b> = <span color=""orange"">Text</font id=""orange""><br>" & strLE & _
	"<i>Example:</i> <b>[violet]</b>Text<b>[/violet]</b> = <span color=""violet"">Text</font id=""violet""><br>" & strLE & _
	"<i>Example:</i> <b>[yellow]</b>Text<b>[/yellow]</b> = <span color=""yellow"">Text</font id=""yellow""><br>" & strLE & _
	"<i>Example:</i> <b>[green]</b>Text<b>[/green]</b> = <span color=""green"">Text</font id=""green""><br>" & strLE & _
	"<i>Example:</i> <b>[gold]</b>Text<b>[/gold]</b> = <span color=""gold"">Text</font id=""gold""><br>" & strLE & _
	"<i>Example:</i> <b>[white]</b>Text<b>[/white]</b> = <span color=""white"">Text</font id=""white""><br>" & strLE & _
	"<i>Example:</i> <b>[purple]</b>Text<b>[/purple]</b> = <span color=""purple"">Text</font id=""purple"">" & strLE & _
	"</p>" & strLE & _
	"<p>&nbsp; </p>" & strLE & _
	"<p><b>Headings:</b><br>" & strLE & _
	"Enclose your text with [h<i>number</i>] and [/h<i>n</i>]<br>" & strLE & _
	"<table>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h1]</b>Text<b>[/h1]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h1>Text</h1>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h2]</b>Text<b>[/h2]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h2>Text</h2>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h3]</b>Text<b>[/h3]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h3>Text</h3>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h4]</b>Text<b>[/h4]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h4>Text</h4>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h5]</b>Text<b>[/h5]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h5>Text</h5>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<i>Example:</i> <b>[h6]</b>Text<b>[/h6]</b> =" & strLE & _
	"</span></td>" & strLE & _
	"<td><span class=""dff dfs ffc"">" & strLE & _
	"<h6>Text</h6>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</p>" & strLE & _
	"<p>&nbsp; </p>" & strLE & _
	"<p><b>Font Sizes:</b><br>" & strLE & _
	"<i>Example:</i> <b>[size=1]</b>Text<b>[/size=1]</b> = <span class=""size=""1"">Text</font id=""size1""><br>" & strLE & _
	"<i>Example:</i> <b>[size=2]</b>Text<b>[/size=2]</b> = <span class=""size=""2"">Text</font id=""size2""><br>" & strLE & _
	"<i>Example:</i> <b>[size=3]</b>Text<b>[/size=3]</b> = <span class=""size=""3"">Text</font id=""size3""><br>" & strLE & _
	"<i>Example:</i> <b>[size=4]</b>Text<b>[/size=4]</b> = <span class=""size=""4"">Text</font id=""size4""><br>" & strLE & _
	"<i>Example:</i> <b>[size=5]</b>Text<b>[/size=5]</b> = <span class=""size=""5"">Text</font id=""size5""><br>" & strLE & _
	"<i>Example:</i> <b>[size=6]</b>Text<b>[/size=6]</b> = <span class=""size=""6"">Text</font id=""size6"">" & strLE & _
	"</p>" & strLE & _
	"<p>&nbsp; </p>" & strLE & _
	"<p><b>Bulleted List:</b> <b>[list]</b> and <b>[/list]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & strLE & _
	"<p><b>Ordered Alpha List:</b> <b>[list=a]</b> and <b>[/list=a]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & strLE & _
	"<p><b>Ordered Number List:</b> <b>[list=1]</b> and <b>[/list=1]</b>, and items in list with <b>[*]</b> and <b>[/*]</b>.</p>" & strLE & _
	"<p><b>Code:</b> Enclose your text with <b>[code]</b> and <b>[/code]</b>.</p>" & strLE & _
	"<p><b>Quote:</b> Enclose your text with <b>[quote]</b> and <b>[/quote]</b>.</p>" & strLE
if (strIMGInPosts = "1") then
	Response.Write "<p><b>Images:</b> Enclose the address with one of the following:<ul><li><b>[img]</b> and <b>[/img]</b></li>" & strLE & _
		"<li><b>[img=right]</b> and <b>[/img=right]</b></li>" & strLE & _
		"<li><b>[img=left]</b> and <b>[/img=left]</b></li></ul></p>" & strLE
end if
Response.Write "</blockquote></span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
WriteFooterShort
Response.End
%>
