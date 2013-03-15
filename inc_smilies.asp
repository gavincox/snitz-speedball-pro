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

Response.Write "<script type=""text/javascript"">" & strLE & _
	"<!-- " & vbNewLine & _
    "function insertsmilie(smilieface) {" & strLE & _
	"AddText(smilieface);" & strLE & _
	"}" & strLE & _
	"// -->" & strLE & _
	"</script>" & strLE & _
	"<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""2"">" & strLE & _
	"<tr class=""c"">" & strLE & _
	"<td class=""c"" colspan=""4""><a name=""smilies""></a><span class=""dff ffs""><b>Smilies</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vam c"">" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:)]')"">" & getCurrentIcon(strIconSmile,"Smile [:)]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:D]')"">" & getCurrentIcon(strIconSmileBig,"Big Smile [:D]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[8D]')"">" & getCurrentIcon(strIconSmileCool,"Cool [8D]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:I]')"">" & getCurrentIcon(strIconSmileBlush,"Blush [:I]","") & "</a></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vam c"">" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:p]')"">" & getCurrentIcon(strIconSmileTongue,"Tongue [:P]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[}:)]')"">" & getCurrentIcon(strIconSmileEvil,"Evil [):]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[;)]')"">" & getCurrentIcon(strIconSmileWink,"Wink [;)]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:o)]')"">" & getCurrentIcon(strIconSmileClown,"Clown [:o)]","") & "</a></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vam c"">" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[B)]')"">" & getCurrentIcon(strIconSmileBlackeye,"Black Eye [B)]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[8]')"">" & getCurrentIcon(strIconSmile8ball,"Eight Ball [8]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:(]')"">" & getCurrentIcon(strIconSmileSad,"Frown [:(]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[8)]')"">" & getCurrentIcon(strIconSmileShy,"Shy [8)]","") & "</a></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vam c"">" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:0]')"">" & getCurrentIcon(strIconSmileShock,"Shocked [:0]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:(!]')"">" & getCurrentIcon(strIconSmileAngry,"Angry [:(!]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[xx(]')"">" & getCurrentIcon(strIconSmileDead,"Dead [xx(]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[|)]')"">" & getCurrentIcon(strIconSmileSleepy,"Sleepy [|)]","") & "</a></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vam c"">" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[:X]')"">" & getCurrentIcon(strIconSmileKisses,"Kisses [:X]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[^]')"">" & getCurrentIcon(strIconSmileApprove,"Approve [^]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[V]')"">" & getCurrentIcon(strIconSmileDisapprove,"Disapprove [V]","") & "</a></td>" & strLE & _
	"<td><a href=""Javascript:insertsmilie('[?]')"">" & getCurrentIcon(strIconSmileQuestion,"Question [?]","") & "</a></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
%>
