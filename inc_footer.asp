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

'Response.Write "</td>" & strLE & _
'	"</tr>" & strLE & _
'	"</table>" & strLE & _
'	"</main>" & strLE & _
Response.Write strLE & _

	"<footer role=""contentinfo"">" & strLE & _
	"<div id=""footer"">" & strLE & _
		"<div class=""fttitle"">" & chkString(strForumTitle,"pagetitle") & "</div>" & strLE & _
		"<div class=""ftcopyright"">&copy; " & strCopyright & "&nbsp;&nbsp;<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam""") & "</a></div>" & strLE & _
	"</div>" & strLE & _
	"<!-- /footer -->" & strLE & _
	"<div id=""subfooter"">" & strLE

if strShowTimer = "1" then
	Response.Write "<div class=""ftgenerated"">" & chkString(replace(strTimerPhrase, "[TIMER]", abs(round(StopTimer(1), 2)), 1, -1, 1),"display") & "</div>" & strLE
end if

Response.Write "<div class=""ftpoweredby"">"
'## START - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
Response.Write "<a href=""http://forum.snitz.com"" target=""_blank"" title=""Powered By: Snitz Forums 2000"">"
if strShowImagePoweredBy = "1" then
	Response.Write getCurrentIcon("logo_powered_by.gif||","Snitz Forums 2000","")
else
	Response.Write "Snitz Forums 2000"
end if
Response.Write "</a>" & strLE & _
	"</div>" & strLE
'## END   - REMOVAL, MODIFICATION OR CIRCUMVENTING THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

Response.Write "</div>" & strLE & _
	"<!-- /subfooter -->" & strLE & _
	"</footer>" & strLE & _
	"<script src=""js/common.js""></script>" & strLE & _
	"</body>" & strLE & _
	"</html>" & strLE
my_Conn.Close
set my_Conn = nothing
%>
