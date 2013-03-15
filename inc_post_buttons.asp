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

Response.Write "<tr>" & strLE & _
		"<td class=""putc vat r"">" & strLE & _
		"<span class=""dff dfs""><b>Format Mode:</b></span></td>" & strLE & _
		"<td class=""putc l"">" & strLE & _
		"<select name=""mode"" onChange=""thelp(this.options[this.selectedIndex].value)"">" & strLE & _
		"                	<option selected value=""0"">Basic&nbsp;</option>" & strLE & _
		"                	<option value=""1"">Help&nbsp;</option>" & strLE & _
		"                	<option value=""2"">Prompt&nbsp;</option>" & strLE & _
		"</select>" & strLE & _
		"<a href=""JavaScript:openWindowHelp('pop_help.asp?mode=post#mode')"">" & getCurrentIcon(strIconSmileQuestion,"Help","class=""vam""") & "</a></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""putc vat r"" rowspan=""2"">" & strLE & _
		"<span class=""dff dfs""><b>Format:</b></span></td>" & strLE & _
		"<td class=""putc l"">" & strLE & _
		"<a href=""Javascript:bold();"">" & getCurrentIcon(strIconEditorBold,"Bold","align=""top""") & "</a>" & _
		"<a href=""Javascript:italicize();"">" & getCurrentIcon(strIconEditorItalicize,"Italicized","align=""top""") & "</a>" & _
		"<a href=""Javascript:underline();"">" & getCurrentIcon(strIconEditorUnderline,"Underline","align=""top""") & "</a>" & _
		"<a href=""Javascript:strike();"">" & getCurrentIcon(strIconEditorStrike,"Strikethrough","align=""top""") & "</a>" & strLE & _
		"<a href=""Javascript:left();"">" & getCurrentIcon(strIconEditorLeft,"Align Left","align=""top""") & "</a>" & _
		"<a href=""Javascript:center();"">" & getCurrentIcon(strIconEditorCenter,"Centered","align=""top""") & "</a>" & _
		"<a href=""Javascript:right();"">" & getCurrentIcon(strIconEditorRight,"Align Right","align=""top""") & "</a>" & strLE & _
		"<a href=""Javascript:hr();"">" & getCurrentIcon(strIconEditorHR,"Horizontal Rule","align=""top""") & "</a>" & _
		"<a href=""Javascript:hyperlink();"">" & getCurrentIcon(strIconEditorUrl,"Insert Hyperlink","align=""top""") & "</a>" & _
		"<a href=""Javascript:email();"">" & getCurrentIcon(strIconEditorEmail,"Insert Email","align=""top""") & "</a>"
if strIMGInPosts = "1" then
	Response.Write "<a href=""Javascript:image();"">" & getCurrentIcon(strIconEditorImage,"Insert Image","align=""top""") & "</a>" & strLE
end if
Response.Write "<a href=""Javascript:showcode();"">" & getCurrentIcon(strIconEditorCode,"Insert Code","align=""top""") & "</a>" & _
		"<a href=""Javascript:quote();"">" & getCurrentIcon(strIconEditorQuote,"Insert Quote","align=""top""") & "</a>" & _
		"<a href=""Javascript:list();"">" & getCurrentIcon(strIconEditorList,"Insert List","align=""top""") & "</a>" & strLE
if lcase(strIcons) = "1" and strShowSmiliesTable = "0" then
	Response.Write "<a href=""JavaScript:openWindow2('pop_icon_legend.asp')"">" & getCurrentIcon(strIconEditorSmilie,"Insert Smilie","align=""top""") & "</a>" & strLE
end if
Response.Write "</td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""putc l"">" & strLE & _
		"<span class=""dff dfs"">" & strLE & _
		"<select name=""Font"" onChange=""showfont(this.options[this.selectedIndex].value)"">" & strLE & _
		"                	<option value="""" selected>Font</option>" & strLE & _
		"                	<option value=""Andale Mono"">Andale Mono</option>" & strLE & _
		"                	<option value=""Arial"">Arial</option>" & strLE & _
		"                	<option value=""Arial Black"">Arial Black</option>" & strLE & _
		"                	<option value=""Book Antiqua"">Book Antiqua</option>" & strLE & _
		"                	<option value=""Century Gothic"">Century Gothic</option>" & strLE & _
		"                	<option value=""Comic Sans MS"">Comic Sans MS</option>" & strLE & _
		"                	<option value=""Courier New"">Courier New</option>" & strLE & _
		"                	<option value=""Georgia"">Georgia</option>" & strLE & _
		"                	<option value=""Impact"">Impact</option>" & strLE & _
		"                	<option value=""Lucida Console"">Lucida Console</option>" & strLE & _
		"                	<option value=""Script MT Bold"">Script MT Bold</option>" & strLE & _
		"                	<option value=""Stencil"">Stencil</option>" & strLE & _
		"                	<option value=""Tahoma"">Tahoma</option>" & strLE & _
		"                	<option value=""Times New Roman"">Times New Roman</option>" & strLE & _
		"                	<option value=""Trebuchet MS"">Trebuchet MS</option>" & strLE & _
		"                	<option value=""Verdana"">Verdana</option>" & strLE & _
		"</select>&nbsp;" & strLE & _
		"<select name=""Size"" onChange=""showsize(this.options[this.selectedIndex].value)"">" & strLE & _
		"                	<option value="""" selected>Size</option>" & strLE & _
		"                	<option value=""1"">1</option>" & strLE & _
		"                	<option value=""2"">2</option>" & strLE & _
		"                	<option value=""3"">3</option>" & strLE & _
		"                	<option value=""4"">4</option>" & strLE & _
		"                	<option value=""5"">5</option>" & strLE & _
		"                	<option value=""6"">6</option>" & strLE & _
		"</select>&nbsp;" & strLE & _
		"<select name=""Color"" onChange=""showcolor(this.options[this.selectedIndex].value)"">" & strLE & _
		"                	<option value="""" selected>Color</option>" & strLE & _
		"                	<option style=""color:black"" value=""black"">Black</option>" & strLE & _
		"                	<option style=""color:red"" value=""red"">Red</option>" & strLE & _
		"                	<option style=""color:yellow"" value=""yellow"">Yellow</option>" & strLE & _
		"                	<option style=""color:pink"" value=""pink"">Pink</option>" & strLE & _
		"                	<option style=""color:green"" value=""green"">Green</option>" & strLE & _
		"                	<option style=""color:orange"" value=""orange"">Orange</option>" & strLE & _
		"                	<option style=""color:purple"" value=""purple"">Purple</option>" & strLE & _
		"                	<option style=""color:blue"" value=""blue"">Blue</option>" & strLE & _
		"                	<option style=""color:beige"" value=""beige"">Beige</option>" & strLE & _
		"                	<option style=""color:brown"" value=""brown"">Brown</option>" & strLE & _
		"                	<option style=""color:teal"" value=""teal"">Teal</option>" & strLE & _
		"                	<option style=""color:navy"" value=""navy"">Navy</option>" & strLE & _
		"                	<option style=""color:maroon"" value=""maroon"">Maroon</option>" & strLE & _
		"                	<option style=""color:limegreen"" value=""limegreen"">LimeGreen</option>" & strLE & _
		"</select></span></td>" & strLE & _
		"</tr>" & strLE
%>
