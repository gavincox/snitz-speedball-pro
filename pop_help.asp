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
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE
select case Request.QueryString("mode")
	case "post"
		'### Format Mode Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""mode""></a><span class=""dff dfs cfc""><b>What is Format Mode used for?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc"">" & strLE & _
			"<li><span class=""dff dfs ffc""><b>Basic:</b>&nbsp;Adds the Forum Code tags to the Message Box</span></li>" & strLE & _
			"<li><span class=""dff dfs ffc""><b>Prompt:</b>&nbsp;Opens a javascript box for you to put your text in</span></li>" & strLE & _
			"<li><span class=""dff dfs ffc""><b>Help:</b>&nbsp;Displays an alert box with a description of the button</span></li>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE
	case "options"
		'### Category Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""category""></a><span class=""dff dfs cfc""><b>Category</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Select the category you would like to place your new forum/url in or move your existing forum/url to.<br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Address Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""address""></a><span class=""dff dfs cfc""><b>Address</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Enter the url to the site you want to create a web link to. Make sure to prefix the Address with <b>http://</b>, <b>https://</b> or <b>file:///</b>.<br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Default Days Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""defaultdays""></a><span class=""dff dfs cfc""><b>Default Days</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"This option allows you to select the default amount of days of topics that is displayed on the Forum page (forum.asp), if the member has not selected an option.<br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Forum Count Member Posts Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""forumcntmposts""></a><span class=""dff dfs cfc""><b>Increase Post Count</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"This option allows you to select whether a Member's Post Count will increase when they make a post in this forum.<br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Moderators Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""moderators""></a><span class=""dff dfs cfc""><b>Moderators</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Here you will be able to select which moderator/s you wish to moderate this forum. Use the buttons to move selected moderators from one list to the other and/or move the whole list of moderators.<br><br>" & strLE & _
			"<b>Available:</b>&nbsp;This list contains the usernames of all the moderators on your forum that are available . If only the Admin account is shown, it means you haven't selected any accounts as moderators.<br><br>" & strLE & _
			"<b>Selected:</b>&nbsp;This list contains the usernames of all the moderators who you have chosen to moderate this forum.</span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Subscriptions Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""subscription""></a><span class=""dff dfs cfc""><b>Subscriptions</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Select the highest level of Subscriptions you would like for this Category/Forum.<br><br>" & strLE & _
			"<b>Category Subscriptions Allowed:</b>&nbsp;This allows registered members to subscribe to the entire category, which will notify them of any posts made within any topic, within any forum, within the category.<br><br>"  & strLE & _
			"<b>Forum Subscriptions Allowed:</b>&nbsp;This allows registered members to subscribe to the entire forum, which will notify them of any posts made within any topic, within the forum.<br><br>"  & strLE & _
			"<b>Topic Subscriptions Allowed:</b>&nbsp;This allows registered members to subscribe to individual topics only, which will notify them of any post made within the topic.<br><br></span>"  & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Moderation Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""moderation""></a><span class=""dff dfs cfc""><b>Moderation</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Select the type of Moderation you want for this forum.<br><br>" & strLE & _
			"<b>All Posts Moderated:</b>&nbsp;This option allows you to moderate all posts made to the forum. Every new topic or reply that is made in the forum will be put on hold until an admin/moderator approves the post.<br><br>" & strLE & _
			"<b>Original Posts Only Moderated:</b>&nbsp;This option allows you to moderate only the new topics that are posted to the forum. Replies are not moderated.<br><br>" & strLE & _
			"<b>Replies Only Moderated:</b>&nbsp;This option allows you to moderate only the replies that are posted to the forum. New topics are not moderated.<br><br>" & strLE & _
			"<i>Note: Admins and Moderators posts are <b>not</b> moderated.</i><br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Authorization Type Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""authtype""></a><span class=""dff dfs cfc""><b>Authorization Type</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"The Authorization Type allows you to place restrictions on who is allowed to access the forum. A description of each type is outlined below:<br><br>" & strLE & _
			"<b>All Visitors:</b>&nbsp;This allows all members (including unregistered members) to access the forum. This is selected by default.<br><br>" & strLE & _
			"<b>Members Only:</b>&nbsp;This allows only registered members to access the forum. Unregistered members are not allowed access.<br><br>" & strLE & _
			"<b>Members Only (Hidden):</b>&nbsp;This allows only registered members to access the forum. The forum will be hidden to unregistered members and they are not allowed access.<br><br>" & strLE & _
			"<b>Password Protected:</b>&nbsp;This allows you to set a password on the forum. All members (including unregistered members) will be asked for a password before giving them access. Once they supply the correct password, they won't be asked for the password again.<br><br>" & strLE & _
			"<b>Members Only & Password Protected:</b>&nbsp;This allows all registered members to access the forum <b>OR</b> if they are not a registered member, they will be asked for the password. Once they enter the correct password, they won't be asked for the password again.<br><br>" & strLE & _
			"<b>Allowed Member List & Password Protected:</b>&nbsp;This allows the members that you select from the Available Members List, to access the forum <b>OR</b> if they are not in the Selected Members List, they will be asked for the password that you set. Once they enter the correct password, they won't be asked for the password again.<br><br>" & strLE & _
			"<b>Allowed Member List:</b>&nbsp;This allows only the members that you select from the Available Members List, to access the forum. All other members (including unregistered members) are not granted access.<br><br>" & strLE & _
			"<b>Allowed Member List (Hidden):</b>&nbsp;This allows only the members that you select from the Available Members List, to access the forum. The forum is hidden from all other members (including unregistered members) who are not on the Selected Members List.<br><br>" & strLE & _
			"<i>Note: Administrators have access to all forums, despite what Authorization is set.</i><br></span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
		'### Allowed Members List Help
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""memberlist""></a><span class=""dff dfs cfc""><b>Allowed Member List</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs ffc"">" & strLE & _
			"Here you will be able to select which registered member or members will be able to have access to the forum. Use the buttons to move selected members from one list to the other and/or move the whole list of members. This option is only valid when you have <b>Allowed Member List</b>, <b>Allowed Member List & Password Protected</b> or <b>Allowed Member List (Hidden)</b> selected for the Auth Type option.<br><br>" & strLE & _
			"<b>Available:</b>&nbsp;This list contains the usernames of all registered members on your forum that are available.<br><br>" & strLE & _
			"<b>Selected:</b>&nbsp;This list contains the usernames of the members who you have selected to access the forum.</span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE
	case else
		'### No Mode Selected
		Response.Write "<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff dfs dfc""><center><b>No mode selected</b></center></span></td>" & strLE & _
			"</tr>" & strLE
end select
		Response.Write "</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE
Call WriteFooterShort()
%>