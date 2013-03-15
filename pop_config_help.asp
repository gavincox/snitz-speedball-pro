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
	case "system"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""strConnString""></a><span class=""dff dfs cfc""><b>How do I configure the strConnString?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc"">" & strLE & _
			"<li><span class=""dff ffs ffc""><b>DSN:</b></span><br>" & strLE & _
			"<span class=""dff ffs ffc"">snitz_forum</span></li>" & strLE & _
			"<li><span class=""dff ffs ffc""><b>MS Access DSN-less:</b></span><br>" & strLE & _
			"<span class=""dff ffs ffc"">strConnString = &quot;DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=c:\www\snitz.com\db\snitz_forum.mdb&quot;</span></li>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""tableprefix""></a><span class=""dff dfs cfc""><b>What's Table Name Prefix?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Table Name Prefix is used if you have multiple versions of the forum running in the same database. This way you can name the tables differently and still use one user to connect. (eg. FORUM_ and FORUM2_)" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""forumtitle""></a><span class=""dff dfs cfc""><b>What's Forum Title?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Forum Title is the title that shows up in the upper right hand corner of the forum. It is also used in e-mails to show where the e-mail came from when posting replies are sent and when new users register." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""copyright""></a><span class=""dff dfs cfc""><b>What's Forum Copyright?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This copyright statements location is basically saying that any topics or replies that are posted are copyrighted material of your organization. This copyright location also helps to copyright the images of your logo and any other material that may be posted on forum pages; however, it is understood by copyright statements in code and informational pages, that the forum code itself is still copyright &copy; 2000 Snitz Communications.<br><br><b><span class=""hlfc"">NOTE:</b>  The &copy; will be included automatically.</span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""titleimage""></a><span class=""dff dfs cfc""><b>What's Title Image?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Use a relative path to point to the image you want to show up in the upper left-hand corner of your forum window.<br>" & strLE & _
			"<br>" & strLE & _
			" For example:<br>" & strLE & _
			"<b>bboard_snitz.gif</b><br>" & strLE & _
			" This points to the bboard_snitz.gif graphic in the same directory, whereas the following would point to the root of the web server and up into the base /images/ directory:<br>" & strLE & _
			"<b>../images/bboard_snitz.gif</b>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""homeurl""></a><span class=""dff dfs cfc""><b>What's the Home URL?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" The Home URL is the base address for your website. An example would be:<br>" & strLE & _
			"<b>forum.snitz.com</b><br>" & strLE & _
			"<br>" & strLE & _
			"<span class=""hlfc"">NOTE: Include the full path of the URL whether it begins with <b>http://</b> in front or a relative URL such as <b>../</b>.</span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""forumurl""></a><span class=""dff dfs cfc""><b>What's the Forum URL?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" The Forum URL is the base address for your forum. An example would be:<br>" & strLE & _
			"<b>http://forum.snitz.com/forum</b>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""imagelocation""></a><span class=""dff dfs cfc""><b>What is the Images Location?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Enter the location where your images are located.<br>" & strLE & _
			" If you have not moved the images from their default location, then just leave this field blank.<br><br>" & strLE & _
			" But, if you have created an <b>images</b> directory in your <b>forum</b> directory then enter:<br><br>" & strLE & _
			"<b>images/</b><br><br>in the field." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""AuthType""></a><span class=""dff dfs cfc""><b>Authorization Type?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" You can either select DataBase or NT Domain authorization." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""SetCookieToForum""></a><span class=""dff dfs cfc""><b>Set Cookie To...</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" You can tell your forum to set it's cookie to either the forum, or the base website. You would set it to the forum if you were hosting multiple forums on the same server or the same domain, and they each had different user communities, otherwise you want this feature set to Website and NOT Forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""GfxButtons""></a><span class=""dff dfs cfc""><b>Graphic Buttons?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" By enabling this feature, the forums will use pictures/graphics instead of the default buttons for ""Submit"" and ""Reset"" etc..." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""PoweredBy""></a><span class=""dff dfs cfc""><b>Use Graphic for ""Powered By"" link?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Toggles between using a Graphic Powered By Link, or a Text Powered By Link.  Either way, you must have one or the other..." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ProhibitNewMembers""></a><span class=""dff dfs cfc""><b>Prohibit New Members?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Toggles between allowing or disallowing people to Register on your Forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""RequireReg""></a><span class=""dff dfs cfc""><b>Require Registration?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" When this option in set to <b>On</b>, only registered members who are logged in will be able to view your Forum.  Everyone else will be presented with a login screen." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""UserNameFilter""></a><span class=""dff dfs cfc""><b>UserName Filter?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" When this option in set to <b>On</b>, the names (or names that contain words) that you specify in the UserName Filter configuration will not be available for user's to register with." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "features"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""secureadminmode""></a><span class=""dff dfs cfc""><b>Secure Admin Mode?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs hlfc"">" & strLE & _
			"<b>WARNING: Only turn Secure Admin off if you absolutely need to. If this option is turned off, anyone can change your forum's configuration!</b>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""allownoncookies""></a><span class=""dff dfs cfc""><b>Why would I want Non-Cookie Mode on?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" If your user base does not use cookies, then you would want to turn this function ""ON"". WARNING: all your admin functions will be visible to all users if this function is ""ON""." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""IPLogging""></a><span class=""dff dfs cfc""><b>What is IP Logging?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" IP Logging will record in the database the IP address of the person who posted a new Topic or Reply. A moderator or administrator then could view the IP by clicking on an icon above the post in the topic." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""FloodCheck""></a><span class=""dff dfs cfc""><b>What is Flood Control?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" With Flood Control enabled, normal users will have to wait the specified amount of time between posts before they can post again." & strLE & _
			"<br><br>Admins and Moderators are not affected by this limitation." & strLE & _
			"<br><br>You can choose 30 seconds, 60 seconds, 90 seconds or 120 seconds." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""privateforums""></a><span class=""dff dfs cfc""><b>What are Private Forums?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Private Forums enable you to only allow certain members to see that the forum exists. If it's only password protected, everyone can see that it exists, however, they are prompted for a password to get in." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""groupcategories""></a><span class=""dff dfs cfc""><b>What are Group Categories?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Group Categories enable you to ""group"" Categories together into ""Groups"" to better organize how Categories are displayed on your forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Subscription""></a><span class=""dff dfs cfc""><b>What is Highest level of Subscription for?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allows you to set the Highest Level of Subscription that can be used on the Forum.  You will also need to set the individual level in each of your Categories and Forums." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""badwordfilter""></a><span class=""dff dfs cfc""><b>Bad Word Filter?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Screen out words you and your guests would find offensive.<br><br>Bad Words can by configured via the Bad Word Configuration option in the Admin Options." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Moderation""></a><span class=""dff dfs cfc""><b>What does Allow Topic Moderation do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" When enabled, this feature allows the Administrator or the Moderator to ""Approve"", ""Hold"" or ""Delete"" a users post before it is shown to the public." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowModerator""></a><span class=""dff dfs cfc""><b>What does Show Moderators do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Basically, if this function is on, it shows the name of the moderator beside the forum that they moderate on the main default page. If it is off, however, visitors won't see who is moderating the forum they are posting in." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""MoveTopicMode""></a><span class=""dff dfs cfc""><b>Why Restrict Moderators from Moving Posts?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This feature either allows or dis-allows a Moderator of one forum to move topics within their forum to someone else's forum where they do not have moderator rights." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""MoveNotify""></a><span class=""dff dfs cfc""><b>Can I notify the Author if his Topic is moved?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" If enabled, this feature automatically sends an e-mail to the topic author if it is moved." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ArchiveState""></a><span class=""dff dfs cfc""><b>What are Archive Functions?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This toggles whether the icons/links show up for the Archive Functions of this Forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""stats""></a><span class=""dff dfs cfc""><b>What does Show Detailed Statistics do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of detailed statistics (last visited date and time, last post, active topics, newest member) at the bottom of the forum." & strLE & _
			" When turned off, some statistics are displayed at the top of the page." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""JumpLastPost""></a><span class=""dff dfs cfc""><b>What does Show Jump To Last Post Link do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of a Jump To Last Post Link " & getCurrentIcon(strIconLastpost,"","class=""vam""") & " icon on the Default page, Forum page and Active Topics page.  This link will take the user to the last post in that topic." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""showpaging""></a><span class=""dff dfs cfc""><b>What does Show Quick Paging do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Quick Paging is when you have a topic that is more than 1 page, a small graphic and the #'s will be show next to the topic title so you can go straight to page 2 or 3, etc..." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""pagenumbersize""></a><span class=""dff dfs cfc""><b>What is Pagenumbers per row for?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This is now only used for the Topic Paging, it limits the amount of pages shown in each row when a topic is more than one page long." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""StickyTopic""></a><span class=""dff dfs cfc""><b>What does Allow Sticky Topics do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the ability of an Admin or Moderator to ""Stick"" a post at the top of the Topics List.  While this Topic is ""Sticky"", it will remain at the top of the list." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""editedbydate""></a><span class=""dff dfs cfc""><b>What would Edited By on Date do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" When a post is edited, there is an appending to the end of the post that says when and by whom the post was edited. Turning this function off would make it so that the footer would not be placed on the end of the post." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowTopicNav""></a><span class=""dff dfs cfc""><b>What does Show Prev / Next Topic do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of previous topic " & getCurrentIcon(strIconGoLeft,"","class=""vam""") & " and next topic " & getCurrentIcon(strIconGoRight,"","class=""vam""") & " icons on the topics page." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowSendToFriend""></a><span class=""dff dfs cfc""><b>What does Show Send to a Friend Link do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of a Send Topic to a Friend Link that is shown when viewing a topic..  This link will allow a user to e-mail a topic to a friend.  E-mail functions must be on for this link to show up." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowPrinterFriendly""></a><span class=""dff dfs cfc""><b>What does Show Printer Friendly Link do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of a Printer Friendly link that is shown when viewing a topic.  This link will popup a window with the topic and any replies that are shown in a format that is easier to print." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""hottopics""></a><span class=""dff dfs cfc""><b>What are Hot Topics?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Hot Topics change the topic folder icon in the Forum view from a normal folder to a flaming folder to let people know that your minimum number of posts has been met to categorize this topic as one that is seeing a lot of action." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""pagesize""></a><span class=""dff dfs cfc""><b>What is Items per page for?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This is the maximum amount of items shown on each page. Once the amount of items on the page reaches this amount, a dropdown box will be shown where you can select other pages." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""AllowHTML""></a><span class=""dff dfs cfc""><b>Why would I allow HTML?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" By allowing HTML you are opening up a whole big can of worms. You may wish to allow HTML in a controlled INTRANET environment,though. It is not recommended to be used on the INTERNET as anyone can post anything without your being able to screen it. IE Pornographic pictures, JavaScript that messes up your pages, etc..." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""AllowForumCode""></a><span class=""dff dfs cfc""><b>Enable Forum Code?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" By turning off Forum Code, you can allow users to mark up their posts with safe codes." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""imginposts""></a><span class=""dff dfs cfc""><b>Why enable Images in Posts?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allows users to place images into their Posts. However, you should be aware that this feature would allow anyone to post ANY image in your forums. This may lead to broken links and potentially objectionable material being displayed." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""icons""></a><span class=""dff dfs cfc""><b>What do Icons do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow users to post smiley faces and other icons allowed by the Forums within the body of their posts!" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""signatures""></a><span class=""dff dfs cfc""><b>Why enable Signatures?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allows users to set a ""Signature"" into their Posts. The same concerns mentioned for Images in Posts applies here as well." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""dsignatures""></a><span class=""dff dfs cfc""><b>Why enable Dynamic Signatures?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" First, you must have Signatures enabled to use Dynamic Signatures.  With Dynamic Signatures enabled, the users signature is not added to the post until it is viewed, so if a person changes their signature, that change will apply to all posts made by that user.  But, this will only apply to posts made while Dyanmic Signatures are enabled.  Any signature that is already in a post won't be updated." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowFormatButtons""></a><span class=""dff dfs cfc""><b>Show Format Buttons?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This turns off or on the Format Section on the screen where your users post new topics/reply to existing topics.<br><br><span class=""hlfc"">Note:</span>&nbsp;You must also have Forum Code enabled on your forum to use this feature." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowSmiliesTable""></a><span class=""dff dfs cfc""><b>Show Smilies Table?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allows users to insert smilies in their posts by clicking on the smilie in a small table shown to them in the post screen.<br><br><span class=""hlfc"">Note:</span>&nbsp;You must also have Icons enabled on your forum to use this feature." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowQuickReply""></a><span class=""dff dfs cfc""><b>Show Quick Reply?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allows users to reply to a topic via a reply box at the bottom of the page when viewing a topic." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""timer""></a><span class=""dff dfs cfc""><b>What does Show Timer do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off the display of the time it took (in seconds) to generate/display the current page.  This time is shown in the footer of every (non popup) page." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""timerphrase""></a><span class=""dff dfs cfc""><b>What is Timer Phrase?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This is what will display in the footer of every (non popup) page.  The phrase must contain the <b>[TIMER]</b> placeholder.  This is where the actual time will be in the phrase (it's dynamically inserted when the page is created)." & strLE & _
			"<br><br><b><span class=""hlfc"">Show Timer must be enabled for this to be used.</span></b>" & strLE & _
			"<br><br>The default is:  <b>This page was generated in [TIMER] seconds.</b>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "members"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""FullName""></a><span class=""dff dfs cfc""><b>What is Fullname For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Full Name (First Name and Last Name), to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Picture""></a><span class=""dff dfs cfc""><b>What is Picture For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter a link to a Picture of themselves, to be viewed in their profile.<br><br>As Admin, you should review the picture in each user's profile from time to time to be sure that the Picture linked to is appropriate for your Forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""RecentTopics""></a><span class=""dff dfs cfc""><b>What is Recent Topics For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" When Recent Topics is enabled, a list of the last 10 Topics posted to by a user will be shown in their Profile.<br><br>This includes New Topics and replies to existing topics." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Sex""></a><span class=""dff dfs cfc""><b>What is Sex For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Sex (either Male or Female), to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Age""></a><span class=""dff dfs cfc""><b>What is Age For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their age, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""AgeDOB""></a><span class=""dff dfs cfc""><b>What is Birth Date For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Birth Date, from which their Age will be calculated and displayed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""MinAge""></a><span class=""dff dfs cfc""><b>What is Minimum Age for?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Prevent users under the age you specify here from registering. The default is 13 for COPPA compliancy but you can change it to anything you want. To turn this feature off completely, set the minimum age to 0." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""City""></a><span class=""dff dfs cfc""><b>What is City For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their City, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""State""></a><span class=""dff dfs cfc""><b>What is State For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their State, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Country""></a><span class=""dff dfs cfc""><b>What is Country For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to choose their Country, to be viewed in their profile and in each Topic or Reply they post." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""aim""></a><span class=""dff dfs cfc""><b>What is the AIM Option For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off features that allow users to enter their AIM username... then for other users to send them messages and/or add them to their buddy list." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""icq""></a><span class=""dff dfs cfc""><b>What is the ICQ Option For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off features that allow users to enter their ICQ number... then for other users to send them ICQ messages and/or see if they are online." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""msn""></a><span class=""dff dfs cfc""><b>What is the MSN Option For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off features that allow users to enter their MSN username... then for other users to view their MSN Username." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""yahoo""></a><span class=""dff dfs cfc""><b>What is the YAHOO Option For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Turns On/Off features that allow users to enter their YAHOO username... then for other users to send them messages and/or add them to their buddy list." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Occupation""></a><span class=""dff dfs cfc""><b>What is Occupation For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Occupation, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Homepages""></a><span class=""dff dfs cfc""><b>What is Homepages For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to display their homepage link by their name on each post and in their Profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""FavLinks""></a><span class=""dff dfs cfc""><b>What is Favorite Links For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter 2 of their Favorite Links, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""MStatus""></a><span class=""dff dfs cfc""><b>What is Marital Status For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Marital Status, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Bio""></a><span class=""dff dfs cfc""><b>What is Bio For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Bio, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Hobbies""></a><span class=""dff dfs cfc""><b>What is Hobbies For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Hobbies, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""LNews""></a><span class=""dff dfs cfc""><b>What is Latest News For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Latest News, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""Quote""></a><span class=""dff dfs cfc""><b>What is Quote For?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Allow your users to enter their Quote, to be viewed in their profile." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "ranks"
		arrStarColors = ("Gold|Silver|Bronze|Orange|Red|Purple|Blue|Cyan|Green")
		arrIconStarColors = array(strIconStarGold,strIconStarSilver,strIconStarBronze,strIconStarOrange,strIconStarRed,strIconStarPurple,strIconStarBlue,strIconStarCyan,strIconStarGreen)
		strStarColor = split(arrStarColors, "|")

		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""ShowRank""></a><span class=""dff dfs cfc""><b>Showing Ranks?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			"<ol>" & strLE & _
			"<li>Don't Show Any</li>" & strLE & _
			"<li>Show Rank Only</li>" & strLE & _
			"<li>Show Stars Only</li>" & strLE & _
			"<li>Show Both Stars and Rank</li>" & strLE & _
			"</ol>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""RankColor""></a><span class=""dff dfs cfc""><b>Color of Stars?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" You can change the color of stars that show up for each rank of member. (only when the Stars function is turned on)" & strLE & _
			" Available colors for the stars:<br><br>" & strLE
		for c = 0 to ubound(strStarColor)
			Response.Write " " & getCurrentIcon(arrIconStarColors(c),"","class=""vam""") & "&nbsp;&nbsp;" & strStarColor(c)
			if c <> ubound(strStarColor) then Response.Write("<br>" & strLE) else Response.Write(vbNewLine)
		next
		Response.Write "<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "datetime"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""timetype""></a><span class=""dff dfs cfc""><b>Time Display?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Choose 24Hr to display all times in military (24 hour) format or 12Hr to display all times in 12 hour format appended with an AM or PM depending on whether it's before or after midday. Default is 24 hour." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""TimeAdjust""></a><span class=""dff dfs cfc""><b>Time Adjustment?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Enter either a positive or negative integer value between +12 and 0 and -12. This may come in handy if you are located in one part of the world, and your server is in another, and you need the time displayed in the forum to be converted to a local time for you! (Default value is 0, meaning no adjustment)" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr> " & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""datetype""></a><span class=""dff dfs cfc""><b>Date Display?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Choose the format you wish all dates to be displayed in. Default is 12/31/2000 (US Short)." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "email"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""email""></a><span class=""dff dfs cfc""><b>What does E-mail do?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Disabling the E-mail function will turn off any features that involve sending mail. If you don't have an SMTP server of any type, you will want to turn this feature off. If you do have an SMTP (mail) server, however, then also select the type of server you have from the dropdown menu." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""mailserver""></a><span class=""dff dfs cfc""><b>What is a Mail Server Address?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" The mail server address is the actual domain name that resolves your mail server. This could be something like:<br>" & strLE & _
			"<b>mail.snitz.com</b><br>" & strLE & _
			" or it could be the same address as the web server:<br>" & strLE & _
			"<b>www.snitz.com</b><br>" & strLE & _
			" Either way, don't put the <b>http://</b> on it.<br>" & strLE & _
			"<br>" & strLE & _
			"<span class=""hlfc""><b>NOTE:</b> If you are using CDONTS as a mail server type, you do not need to fill in this field.</span>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""sender""></a><span class=""dff dfs cfc""><b>Administrator E-mail Address?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This address is referenced by the forums in a couple ways.<br>" & strLE & _
			"<ol>" & strLE & _
			"<li>When mail is sent, it is sent from this user E-mail Account.</li>" & strLE & _
			"<li>This user is also the point of contact given if there is a problem with these forums.</li>" & strLE & _
			"</ol>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""UniqueEmail""></a><span class=""dff dfs cfc""><b>Unique E-mail Address?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Do you want to require each user to have their own E-mail Address?" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""EmailVal""></a><span class=""dff dfs cfc""><b>E-mail Validation?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Do you want to require each user to validate their E-mail Address when they first Register and anytime they change their E-mail Address?<br><br>The user will receive an E-mail with a link in it that will validate that the E-mail Address they entered is a valid E-mail Address." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""EmailFilter""></a><span class=""dff dfs cfc""><b>Filter known spam domains?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This allows you to filter out E-mail addresses from a given domain - like all addresses from @example.com.<br><br>This will prevent people from registering with E-mail addresses at that domain." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""RestrictReg""></a><span class=""dff dfs cfc""><b>Restrict Registration?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This allows you to choose who is able to register on your forum by approving or rejecting their registration.<br><br><b>Note:</b> You must have the E-mail Validation option turned On to use this feature." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""LogonForMail""></a><span class=""dff dfs cfc""><b>Require Logon for sending Mail?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Do you require a user to be logged on before being able to use the <i>Send Topic To a Friend</i> or <i>E-mail Poster</i> options?" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""MaxPostsToEMail""></a><span class=""dff dfs cfc""><b>Number of posts to allow sending e-mail?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" To prevent spammers from registering an account and immediately sending your members e-mails through the forum, set this to the number of posts you want someone to have before they can use the forum's e-mail function.<br><br>Set this to 0 if you want to turn this feature off.<br><br>" & _
			" There is an Admin overide for this feature. If you want to exempt an individual, edit their profile.<br><br><b>Note:</b> this does not affect Admins or Moderators." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""NoMaxPostsToEMail""></a><span class=""dff dfs cfc""><b>Error if they don't have enough posts?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This is the message you want someone to see if they don't have enough posts to send an email.<br><br><b>Note:</b> ""Number of posts to allow sending e-mail"" must be greater than 0." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
	case "colors"
		Response.Write "<tr>" & strLE & _
			"<td class=""ccc""><a name=""fontfacetype""></a><span class=""dff dfs cfc""><b>Font Face Type?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Font Face Type changes the way the text in your forum looks. You may want to change this option to match that of the rest of your web site. Some standards are:" & strLE & _
			"<ul>" & strLE & _
			"<li>Arial (nice, clean, legible font)</li>" & strLE & _
			"<li>Courier (a typewriter font)</li>" & strLE & _
			"<li>Helvetica (another clean, legible font)</li>" & strLE & _
			"<li>Sans Serif (Arial & Helvetica are variants of Sans Serif)</li>" & strLE & _
			"<li>Times New Roman (a book-type font)</li>" & strLE & _
			"<li>Verdana (another clean, legible font) (default)</li>" & strLE & _
			"</ul>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""fontsize""></a><span class=""dff dfs cfc""><b>What does Font Size mean?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			"<ul>" & strLE & _
			"<li>None = Use Browser Default</li>" & strLE & _
			"<li>1 = 8 point font <b>X-Small</b> (default footer size)</li>" & strLE & _
			"<li>2 = 10 point font <b>Small</b> (default font size)</li>" & strLE & _
			"<li>3 = 12 point font <b>Normal</b></li>" & strLE & _
			"<li>4 = 14 point font <b>Large</b> (default header size)</li>" & strLE & _
			"<li>5 = 18 point font <b>X-Large</b></li>" & strLE & _
			"<li>6 = 24 point font <b>XX-Large</b></li>" & strLE & _
			"<li>7 = 36 point font <b>XXX-Large</b></li>" & strLE & _
			"</ul>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""colors""></a><span class=""dff dfs cfc""><b>What colors may I use?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc"">" & strLE & _
			"<p><span class=""dff ffs ffc"">" & strLE & _
			" There are a lot of different colors you can choose from, all of which are listed below:</p>" & strLE & _
			"<blockquote><pre><span class=""dff ffs"">" & strLE & _
			"<span color=""aliceblue"">aliceblue</span>" & strLE & _
			"<span color=""antiquewhite"">antiquewhite</span>" & strLE & _
			"<span color=""aqua"">aqua</span>" & strLE & _
			"<span color=""aquamarine"">aquamarine</span>" & strLE & _
			"<span color=""azure"">azure</span>" & strLE & _
			"<span color=""beige"">beige</span>" & strLE & _
			"<span color=""bisque"">bisque</span>" & strLE & _
			"<span color=""black"">black</span>" & strLE & _
			"<span color=""blanchedalmond"">blanchedalmond</span>" & strLE & _
			"<span color=""blue"">blue</span>" & strLE & _
			"<span color=""blueviolet"">blueviolet</span>" & strLE & _
			"<span color=""brown"">brown</span>" & strLE & _
			"<span color=""burlywood"">burlywood</span>" & strLE & _
			"<span color=""cadetblue"">cadetblue</span>" & strLE & _
			"<span color=""chartreuse"">chartreuse</span>" & strLE & _
			"<span color=""chocolate"">chocolate</span>" & strLE & _
			"<span color=""coral"">coral</span>" & strLE & _
			"<span color=""cornflowerblue"">cornflowerblue</span>" & strLE & _
			"<span color=""cornsilk"">cornsilk</span>" & strLE & _
			"<span color=""cyan"">cyan</span>" & strLE & _
			"<span color=""darkblue"">darkblue</span>" & strLE & _
			"<span color=""darkcyan"">darkcyan</span>" & strLE & _
			"<span color=""darkgoldenrod"">darkgoldenrod</span>" & strLE & _
			"<span color=""darkgray"">darkgray</span>" & strLE & _
			"<span color=""darkgreen"">darkgreen</span>" & strLE & _
			"<span color=""darkkhaki"">darkkhaki</span>" & strLE & _
			"<span color=""darkmagenta"">darkmagenta</span>" & strLE & _
			"<span color=""darkolivegreen"">darkolivegreen</span>" & strLE & _
			"<span color=""darkorange"">darkorange</span>" & strLE & _
			"<span color=""darkorchid"">darkorchid</span>" & strLE & _
			"<span color=""darkred"">darkred</span>" & strLE & _
			"<span color=""darksalmon"">darksalmon</span>" & strLE & _
			"<span color=""darkseagreen"">darkseagreen</span>" & strLE & _
			"<span color=""darkslateblue"">darkslateblue</span>" & strLE & _
			"<span color=""darkslategray"">darkslategray</span>" & strLE & _
			"<span color=""darkturquoise"">darkturquoise</span>" & strLE & _
			"<span color=""darkviolet"">darkviolet</span>" & strLE & _
			"<span color=""deeppink"">deeppink</span>" & strLE & _
			"<span color=""deepskyblue"">deepskyblue</span>" & strLE & _
			"<span color=""dimgray"">dimgray</span>" & strLE & _
			"<span color=""dodgerblue"">dodgerblue</span>" & strLE & _
			"<span color=""firebrick"">firebrick</span>" & strLE & _
			"<span color=""floralwhite"">floralwhite</span>" & strLE & _
			"<span color=""forestgreen"">forestgreen</span>" & strLE & _
			"<span color=""gainsboro"">gainsboro</span>" & strLE & _
			"<span color=""ghostwhite"">ghostwhite</span>" & strLE & _
			"<span color=""gold"">gold</span>" & strLE & _
			"<span color=""goldenrod"">goldenrod</span>" & strLE & _
			"<span color=""gray"">gray</span>" & strLE & _
			"<span color=""green"">green</span>" & strLE & _
			"<span color=""greenyellow"">greenyellow</span>" & strLE & _
			"<span color=""honeydew"">honeydew</span>" & strLE & _
			"<span color=""hotpink"">hotpink</span>" & strLE & _
			"<span color=""indianred"">indianred</span>" & strLE & _
			"<span color=""ivory"">ivory</span>" & strLE & _
			"<span color=""khaki"">khaki</span>" & strLE & _
			"<span color=""lavender"">lavender</span>" & strLE & _
			"<span color=""lavenderblush"">lavenderblush</span>" & strLE & _
			"<span color=""lawngreen"">lawngreen</span>" & strLE & _
			"<span color=""lemonchiffon"">lemonchiffon</span>" & strLE & _
			"<span color=""lightblue"">lightblue</span>" & strLE & _
			"<span color=""lightcoral"">lightcoral</span>" & strLE & _
			"<span color=""lightcyan"">lightcyan</span>" & strLE & _
			"<span color=""lightgoldenrod"">lightgoldenrod</span>" & strLE & _
			"<span color=""lightgoldenrodyellow"">lightgoldenrodyellow</span>" & strLE & _
			"<span color=""lightgray"">lightgray</span>" & strLE & _
			"<span color=""lightgreen"">lightgreen</span>" & strLE & _
			"<span color=""lightpink"">lightpink</span>" & strLE & _
			"<span color=""lightsalmon"">lightsalmon</span>" & strLE & _
			"<span color=""lightseagreen"">lightseagreen</span>" & strLE & _
			"<span color=""lightskyblue"">lightskyblue</span>" & strLE & _
			"<span color=""lightslateblue"">lightslateblue</span>" & strLE & _
			"<span color=""lightslategray"">lightslategray</span>" & strLE & _
			"<span color=""lightsteelblue"">lightsteelblue</span>" & strLE & _
			"<span color=""lightyellow"">lightyellow</span>" & strLE & _
			"<span color=""limegreen"">limegreen</span>" & strLE & _
			"<span color=""linen"">linen</span>" & strLE & _
			"<span color=""magenta"">magenta</span>" & strLE & _
			"<span color=""maroon"">maroon</span>" & strLE & _
			"<span color=""mediumaquamarine"">mediumaquamarine</span>" & strLE & _
			"<span color=""mediumblue"">mediumblue</span>" & strLE & _
			"<span color=""mediumorchid"">mediumorchid</span>" & strLE & _
			"<span color=""mediumpurple"">mediumpurple</span>" & strLE & _
			"<span color=""mediumseagreen"">mediumseagreen</span>" & strLE & _
			"<span color=""mediumslateblue"">mediumslateblue</span>" & strLE & _
			"<span color=""mediumspringgreen"">mediumspringgreen</span>" & strLE & _
			"<span color=""mediumturquoise"">mediumturquoise</span>" & strLE & _
			"<span color=""mediumvioletred"">mediumvioletred</span>" & strLE & _
			"<span color=""midnightblue"">midnightblue</span>" & strLE & _
			"<span color=""mintcream"">mintcream</span>" & strLE & _
			"<span color=""mistyrose"">mistyrose</span>" & strLE & _
			"<span color=""moccasin"">moccasin</span>" & strLE & _
			"<span color=""navajowhite"">navajowhite</span>" & strLE & _
			"<span color=""navy"">navy</span>" & strLE & _
			"<span color=""navyblue"">navyblue</span>" & strLE & _
			"<span color=""oldlace"">oldlace</span>" & strLE & _
			"<span color=""olivedrab"">olivedrab</span>" & strLE & _
			"<span color=""orange"">orange</span>" & strLE & _
			"<span color=""orangered"">orangered</span>" & strLE & _
			"<span color=""orchid"">orchid</span>" & strLE & _
			"<span color=""palegoldenrod"">palegoldenrod</span>" & strLE & _
			"<span color=""palegreen"">palegreen</span>" & strLE & _
			"<span color=""paleturquoise"">paleturquoise</span>" & strLE & _
			"<span color=""palevioletred"">palevioletred</span>" & strLE & _
			"<span color=""papayawhip"">papayawhip</span>" & strLE & _
			"<span color=""peachpuff"">peachpuff</span>" & strLE & _
			"<span color=""peru"">peru</span>" & strLE & _
			"<span color=""pink"">pink</span>" & strLE & _
			"<span color=""plum"">plum</span>" & strLE & _
			"<span color=""powderblue"">powderblue</span>" & strLE & _
			"<span color=""purple"">purple</span>" & strLE & _
			"<span color=""red"">red</span>" & strLE & _
			"<span color=""rosybrown"">rosybrown</span>" & strLE & _
			"<span color=""royalblue"">royalblue</span>" & strLE & _
			"<span color=""saddlebrown"">saddlebrown</span>" & strLE & _
			"<span color=""salmon"">salmon</span>" & strLE & _
			"<span color=""sandybrown"">sandybrown</span>" & strLE & _
			"<span color=""seagreen"">seagreen</span>" & strLE & _
			"<span color=""seashell"">seashell</span>" & strLE & _
			"<span color=""sienna"">sienna</span>" & strLE & _
			"<span color=""skyblue"">skyblue</span>" & strLE & _
			"<span color=""slateblue"">slateblue</span>" & strLE & _
			"<span color=""slategray"">slategray</span>" & strLE & _
			"<span color=""snow"">snow</span>" & strLE & _
			"<span color=""springgreen"">springgreen</span>" & strLE & _
			"<span color=""steelblue"">steelblue</span>" & strLE & _
			"<span color=""tan"">tan</span>" & strLE & _
			"<span color=""thistle"">thistle</span>" & strLE & _
			"<span color=""tomato"">tomato</span>" & strLE & _
			"<span color=""turquoise"">turquoise</span>" & strLE & _
			"<span color=""violet"">violet</span>" & strLE & _
			"<span color=""violetred"">violetred</span>" & strLE & _
			"<span color=""wheat"">wheat</span>" & strLE & _
			"<span color=""white"">white</span>" & strLE & _
			"<span color=""whitesmoke"">whitesmoke</span>" & strLE & _
			"<span color=""yellow"">yellow</span>" & strLE & _
			"<span color=""yellowgreen"">yellowgreen</span>" & strLE & _
			"</span></pre></blockquote>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""fontdecorations""></a><span class=""dff dfs cfc""><b>What are Font Decorations?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			"<ul>" & strLE & _
			"<li>none</li>" & strLE & _
			"<li>blink</li>" & strLE & _
			"<li>line-through</li>" & strLE & _
			"<li>overline</li>" & strLE & _
			"<li>underline</li>" & strLE & _
			"</ul>" & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""pagebgimage""></a><span class=""dff dfs cfc""><b>What is Page Background Image URL?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" Enter the URL to the location of the background image you would like for your forum." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""columnwidth""></a><span class=""dff dfs cfc""><b>How does Column Width Work?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" This sets the width of the column in question. It is not recommended that you change this unless you really know what your doing." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""ccc""><a name=""nowrap""></a><span class=""dff dfs cfc""><b>What is NOWRAP?</b></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""fcc""><span class=""dff ffs ffc"">" & strLE & _
			" NOWRAP prevents the text in a column from auto wrapping. This could be bad if you have people posting long strings of text in the right column (message box), reason being: this would cause an awful long horizontal scroll bar in most cases." & strLE & _
			"<a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","align=""right""") & "</a></span></td>" & strLE & _
			"</tr>" & strLE
end select
Response.Write "</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
Call WriteFooterShort()
%>
