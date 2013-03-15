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

Response.Write "<table width=""100%"">" & strLE & _
		"<tr>" & strLE & _
		"<td><span class=""dff dfs"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;Registration Rules and Policies Agreement</span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE

if strProhibitNewMembers <> "1" then
	Response.Write "<table class=""tc"" width=""100%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""ccc""><span class=""dff dfs cfc""><b>Privacy Statement for " & strForumTitle & "</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
		"<p>If you agree to the terms and conditions stated below, " & _
		"press the &quot;Agree&quot; button. Otherwise, press &quot;Cancel&quot;.</p>" & strLE & _
		"<p>In order to use these forums, users are required to " & _
		"provide a username, password and e-mail address. Neither the Administrators of " & _
		"these forums, or the Moderators participating, are responsible for the privacy " & _
		"practices of any user. Remember that all information that is disclosed in these " & _
		"areas becomes public information and you should exercise caution when deciding " & _
		"to share any of your personal information. Any user who finds material posted by " & _
		"another user objectionable is encouraged to contact us via e-mail. We are " & _
		"authorized by you to remove or modify any data submitted by you to these forums " & _
		"for any reason we feel constitutes a violation of our policies, whether stated, " & _
		"implied or not.</p>" & strLE & _
		"<p>This site may contain links to other web sites and " & _
		"files. We have no control over the content and can not ensure it will not be offensive " & _
		"or objectionable.  We will, however, remove links to material that we feel is inappropriate as we become aware of them.</p>" & strLE & _
		"<p>Cookies must be turned on in your browser to participate " & _
		"as a user in these forums. Cookies are used here to hold your username and " & _
		"password and viewing options, allowing you to login.</p>" & strLE & _
		"<p>By pressing the &quot;Agree&quot; button, you agree that you, the " & _
		"user, are "
	if strMinAge > 0 then
		Response.Write strMinAge
	else
		Response.Write "13"
	end if
	Response.Write " years of age or over. You are fully responsible for any information " & _
		"or file supplied by this user. You also agree that you will not post any " & _
		"copyrighted material that is not owned by yourself or the owners of these " & _
		"forums. In your use of these forums, you agree that you will not post any " & _
		"information which is vulgar, harassing, hateful, threatening, invading of others " & _
		"privacy, sexually oriented, or violates any laws.</p>" & strLE & _
		"<p>If you do agree with the rules and policies stated in " & _
		"this agreement, and meet the criteria stated herein, proceed to press the " & _
		"&quot;Agree&quot; button below, otherwise press &quot;Cancel&quot;.</p>" & strLE & _
		"<hr size=""1"">" & strLE & _
		"<table class=""tc"">" & strLE & _
		"<tbody>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<form action=""register.asp?mode=Register"" id=""form1"" method=""post"" name=""form1"">" & strLE & _
		"<input name=""Refer"" type=""hidden"" value=""" & strReferer & """>" & strLE & _
		"<input name=""policy_accept"" type=""hidden"" value=""true"">" & strLE & _
		"<input name=""Submit"" type=""Submit"" value=""Agree"">" & strLE & _
		"</form>" & strLE & _
		"</td>" & strLE & _
		"<td>" & strLE & _
		"<form action=""JavaScript:history.go(-1)"" id=""form2"" method=""post"" name=""form2"">" & strLE & _
		"<input name=""Submit"" type=""Submit"" value=""Cancel"">" & strLE & _
		"</form>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</tbody>" & strLE & _
		"</table>" & strLE & _
		"<hr size=""1"">" & strLE & _
		"<p>If you have any questions about this privacy statement " & _
		"or the use of these forums, you can contact the forum administrator at: " & _
		"<span class=""smt""><a href=""mailto:" & strSender & """>" & strSender & "</a></span></p>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE & _
		"<br>" & strLE
else
	Response.Write "<br><p class=""c""><span class=""dff hfs hlfc"">Sorry, we are not accepting any new Members at this time.</span></p>" & strLE & _
		"<meta http-equiv=""Refresh"" content=""5; URL=default.asp"">" & strLE & _
		"<p class=""c""><span class=""dff dfs""><a href=""default.asp"">Back To Forum</span></a></p><br>" & strLE
end if
WriteFooter
Response.End
%>
