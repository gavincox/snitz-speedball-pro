<%
'#################################################################################
'## Snitz Forums 2000 v3.4.07
'#################################################################################
'## Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen,
'## Huw Reddick and Richard Kinser
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
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<%
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs w50"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;Frequently Asked Questions</span></td>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<table id=""content"">" & strLE & _
	"<tbody>" & strLE & _
	"<tr>" & strLE & _
	"<th><b>FAQ Table of Contents</b></th>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<ul>" & strLE & _
	"<li class=""smt""><a href=""#register"">Do I have to register?</a></li>" & strLE
if (strIcons = "1") then Response.Write "<li class=""smt""><a href=""#smilies"">How can I use smilies and images?</a></li>" & strLE
Response.Write "<li class=""smt""><a href=""#hyperlink"">Can I add a hyperlink to my messages?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#format"">Can I change the format of my text?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#mods"">What are Moderators?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#cookies"">Are cookies used?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#activetopics"">What are active topics?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#edit"">Can I edit my own posts?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#attach"">Can I attach files?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#search"">Can I search?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#editprofile"">Can I edit my profile?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#signature"">Can I attach my own signature to my posts?</a></li>" & strLE & _
	"<!-- <li class=""smt""><a href=""#announce"">What are announcements?</a></li> -->" & strLE
if (strBadWordFilter = "1") then Response.Write "<li class=""smt""><a href=""#censor"">Are there any censor features?</a></li>" & strLE
if (stremail = "1") then
	Response.Write "<li class=""smt""><a href=""#pw"">What do I do if I forget my Password?</a></li>" & strLE & _
		"<li class=""smt""><a href=""#notify"">Can I be notified by e-mail when there are new posts?</a></li>" & strLE
end if
if (strModeration = "1") then Response.Write "<li class=""smt""><a href=""#moderation"">What does it mean if a forum has Moderation enabled?</a></li>" & strLE
Response.Write "<li class=""smt""><a href=""#COPPA"">What is COPPA?</a></li>" & strLE & _
	"<li class=""smt""><a href=""#GetForum"">Where can I get my own copy of this Forum?</a></li>" & strLE & _
	"<li class=""smt""><a href=""mailto:" & strSender & """>Can't find your answer here? Send us an e-mail.</a></li>" & strLE & _
	"</ul>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""cathd"" id=""register""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><b>Registering</b></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>"
if strProhibitNewMembers = "1" then
	Response.Write "The Administrator has turned off Registration for this forum. Only registered members are able to log in."
elseif strRequireReg = "1" then
	Response.Write "Yes, registration is required. Registration is free and only takes a few minutes. The only required fields are your Username, which may be your real name or a nickname, a Password, and a valid e-mail address.<br><br>"
else
	Response.Write "Registration is not required to view current topics on the Forum; however, if you wish to post a new topic or reply to an existing topic registration is required. Registration is free and only takes a few minutes. The only required fields are your Username, which may be your real name or a nickname, and a valid e-mail address.<br><br>"
end if
if strProhibitNewMembers = "0" then Response.Write "The information you provide during registration is not outsourced or used for any advertising by " & strForumTitle & ".<br><br>If you believe someone is sending you advertisements as a result of the information you provided through your registration, please notify us immediately."
Response.Write "</p></td>" & strLE & _
	"</tr>" & strLE
if (strIcons = "1") then
	strSmileCode = array("[:)]","[:D]","[8D]","[:I]","[:p]","[}:)]","[;)]","[:o)]","[B)]","[8]","[:(]","[8)]","[:0]","[:(!]","[xx(]","[|)]","[:X]","[^]","[V]","[?]")
	strSmileDesc = array("smile","big smile","cool","blush","tongue","evil","wink","clown","black eye","eightball","frown","shy","shocked","angry","dead","sleepy","kisses","approve","disapprove","question")
	strSmileName = array(strIconSmile,strIconSmileBig,strIconSmileCool,strIconSmileBlush,strIconSmileTongue,strIconSmileEvil,strIconSmileWink,strIconSmileClown,strIconSmileBlackeye,strIconSmile8ball,strIconSmileSad,strIconSmileShy,strIconSmileShock,strIconSmileAngry,strIconSmileDead,strIconSmileSleepy,strIconSmileKisses,strIconSmileApprove,strIconSmileDisapprove,strIconSmileQuestion)
	Response.Write "<tr>" & strLE & _
		"<td class=""cathd""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""smilies""></a><b>Smilies</b></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		" You've probably seen others use smilies before in e-mail messages or other bulletin " & strLE & _
		" board posts. Smilies are keyboard characters used to convey an emotion, such as a smile " & strLE & _
		getCurrentIcon(strIconSmile,"","class=""vam""") & " or a frown " & strLE & _
		getCurrentIcon(strIconSmileSad,"","class=""vam""") & ". This bulletin board " & strLE & _
		" automatically converts certain text to a graphical representation when it is " & strLE & _
		" inserted between brackets [].&nbsp; Here are the smilies that are currently " & strLE & _
		" supported by " & strForumTitle & ":<br>" & strLE & _
		"<table class=""tbc w33"">" & strLE & _
		"<tr class=""vat"">" & strLE & _
		"<td>" & strLE & _
		"<table class=""tc"">" & strLE
	for sm = 0 to 9
		Response.Write "<tr>" & strLE & _
			getCurrentIcon(strSmileName(sm),"","class=""vam""") & "</td>" & strLE & _
			strSmileDesc(sm) & "</td>" & strLE & _
			strSmileCode(sm) & "</td>" & strLE & _
			"</tr>" & strLE
	next
	Response.Write "</table>" & strLE & _
		"</td>" & strLE & _
		"<td>" & strLE & _
		"<table class=""tc"">" & strLE
	for sm = 10 to 19
		Response.Write "<tr>" & strLE & _
			getCurrentIcon(strSmileName(sm),"","class=""vam""") & "</td>" & strLE & _
			strSmileDesc(sm) & "</td>" & strLE & _
			strSmileCode(sm) & "</td>" & strLE & _
			"</tr>" & strLE
	next
	Response.Write "</table>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE & _
		"</table></span></p>" & strLE & _
		"</td>" & strLE & _
		"</tr>" & strLE
end if
Response.Write "<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""hyperlink""></a><span class=""dff dfs cfc""><b>Creating a Hyperlink in your message</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>You can easily add a hyperlink to your message.</p><br>" & strLE & _
	"<p>All that you need to do is type the URL (" & strForumURL & "), and it will automatically be converted to a URL (<span class=""smt""><a href=""" & strForumURL & """ target=""_blank"">" & strForumURL & "</a></span>)!</p><br>" & strLE & _
	"<p>The trick here is to make sure you prefix your URL with the <b>http://</b>, <b>https://</b> or <b>file://</b></p><br>" & strLE & _
	"<p>You can also add a mailto link to your message by typing in your e-mail address.<br>" & strLE & _
	"<blockquote>" & strLE & _
	"<i>This Example:</i><br>" & strLE & _
	"<b>" & strSender & "</b><br>" & strLE & _
	"<i>Outputs this:</i><br>" & strLE & _
	"<span class=""smt""><a href=""mailto:" & strSender & """>" & strSender & "</a></span></p>" & strLE & _
	"</blockquote><br>" & strLE & _
	"<p>Another way to add hyperlinks is to use the <b>[url]</b>linkto<b>[/url]</b> tags</p>" & strLE & _
	"<blockquote>" & strLE & _
	"<i>This Example:</i><br>" & strLE & _
	"<b>[url]</b>" & strForumURL & "<b>[/url]</b> takes you home!<br>" & strLE & _
	"<i>Outputs This:</i><br>" & strLE & _
	"<span class=""smt""><a href=""" & strForumURL & """>" & strForumURL & "</a></span> takes you home!" & strLE & _
	"</blockquote></p>" & strLE & _
	"<p>" & strLE & _
	"<p>If you use this tag: <b>[url=&quot;</b>linkto<b>&quot;]</b>description<b>[/url]</b> you can add a description to the link.</p>" & strLE & _
	"<blockquote>" & strLE & _
	"<i>This Example:</i><br>" & strLE & _
	" Take me to <b>[url=&quot;" & strForumURL & "&quot;]</b>" & chkString(strForumTitle,"pagetitle") & "<b>[/url]</b><br>" & strLE & _
	"<i>Outputs This:</i><br>" & strLE & _
	" Take me to <span class=""smt""><a href=""" & strForumURL & """>" & chkString(strForumTitle,"pagetitle") & "</a></span>" & strLE & _
	"</blockquote>" & strLE & _
	"<blockquote>" & strLE & _
	"<i>This Example:</i><br>" & strLE & _
	" If you have a question <b>[url=&quot;" & strSender & "&quot;]</b>E-Mail Me<b>[/url]</b><br>" & strLE & _
	"<i>Outputs This:</i><br>" & strLE & _
	" If you have a question <span class=""smt""><a href=""mailto:" & strSender & """>E-Mail Me</a></span>" & strLE & _
	"</blockquote>" & strLE & _
	"<br>" & strLE
if (strIMGInPosts = "1") then
	Response.Write " You can make clickable images by combining the <b>[url=""</b>linkto<b>""]</b>description<b>[/url]</b> and <b>[img]</b>image_url<b>[/img]</b> tags<br>" & strLE & _
		"<blockquote>" & strLE & _
		"<i>This Example:</i><br>" & strLE & _
		"<b>[url=&quot;" & strForumURL & "&quot;][img]</b>" & strTitleImage & "<b>[/img][/url]</b><br>" & strLE & _
		"<i>Outputs This:</i><br>" & strLE & _
		"<a href=""" & strForumURL & """ target=""_blank"">" & getCurrentIcon(strTitleImage & "||","","") & "</a>" & strLE & _
		"</blockquote>" & strLE & _
		"</p>" & strLE
end if
Response.Write "</span></td>" & strLE & _
	"</tr>" & strLE
if strAllowForumCode = "1" then
	Response.Write "<tr>" & strLE & _
		"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""format""></a><span class=""dff dfs cfc""><b>How to format text with Bold, Italic, Quote, etc...</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		" There are several Forum Codes you may use to change the appearance " & strLE & _
		" of your text.&nbsp; Following is the list of codes currently available:</p>" & strLE & _
		"<blockquote>" & strLE & _
		"<p><b>Bold:</b> Enclose your text with [b] and [/b] .&nbsp; <i>Example:</i> This is <b>[b]</b>bold<b>[/b]</b> text. = This is <b>bold</b> text.</p>" & strLE & _
		"<p><i>Italic:</i> Enclose your text with [i] and [/i] .&nbsp; <i>Example:</i> This is <b>[i]</b>italic<b>[/i]</b> text. = This is <i>italic</i> text.</p>" & strLE & _
		"<p><u>Underline:</u> Enclose your text with [u] and [/u]. <i>Example:</i> This is <b>[u]</b>underline<b>[/u]</b> text. =  This is <u>underline</u> text.</p>" & strLE & _
		"<p><b>Aligning Text Left:</b> Enclose your text with [left] and [/left]" & strLE & _
		"</p>" & strLE & _
		"<p><b>Aligning Text Center:</b> Enclose your text with [center] and [/center]" & strLE & _
		"</p>" & strLE & _
		"<p><b>Aligning Text Right:</b> Enclose your text with [right] and [/right]" & strLE & _
		"</p>" & strLE & _
		"<p><b>Striking Text:</b> Enclose your text with [s] and [/s]<br>" & strLE & _
		" Example: <b>[s]</b>mistake<b>[/s]</b> = <s>mistake</s>" & strLE & _
		"</p>" & strLE & _
		"<p><b>Horizontal Rule:</b> Place a horizontal line in your post with [hr]<br>" & strLE & _
		" Example: <b>[hr]</b> = <hr noshade size=""1"">" & strLE & _
		"</p>" & strLE & _
		"<p>&nbsp; </p>" & strLE & _
		"<p><b>Font Colors:</b><br>" & strLE & _
		" Enclose your text with [<i>fontcolor</i>] and [/<i>fontcolor</i>] <br>" & strLE & _
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
		"<p><b>Headings:</b> Enclose your text with [h<i>number</i>] and [/h<i>n</i>]<br>" & strLE & _
		"<table>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h1]</b>Text<b>[/h1]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
		"<h1>Text</h1>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h2]</b>Text<b>[/h2]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
		"<h2>Text</h2>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h3]</b>Text<b>[/h3]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
		"<h3>Text</h3>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h4]</b>Text<b>[/h4]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
		"<h4>Text</h4>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h5]</b>Text<b>[/h5]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
		"<h5>Text</h5>" & strLE & _
		"</span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<i>Example:</i> <b>[h6]</b>Text<b>[/h6]</b> =" & strLE & _
		"</span></td>" & strLE & _
		"<td>" & strLE & _
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
		"</tr>" & strLE
end if
Response.Write "<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""mods""></a><span class=""dff dfs cfc""><b>Moderators</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>" & strLE & _
	" Moderators control individual forums. They may edit, delete, or prune any posts in their forums."
if (strShowModerators = "1") then Response.Write " If you have a question about a particular forum, you should direct it to your forum moderator."
Response.Write "</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""cookies""></a><span class=""dff dfs cfc""><b>Cookies</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>" & strLE & _
	" These Forums use cookies to store the following information: the last time you logged in, your Username and       your Encrypted Password.  These cookies are stored on your hard drive. Cookies are not used to track your movement or perform any function other than to enhance your use of these forums."
if (strNoCookies = "0") then Response.Write " If you have not enabled cookies in your browser, many of these time-saving features will not work properly. <b>Also, you need to have cookies enabled if you want to enter a private forum or post a topic/reply.</b>"
Response.Write "</p>" & strLE & _
	"<p>You may delete all cookies set by these forums in selecting the &quot;logout&quot; button at the top of any page." & strLE & _
	"</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""activetopics""></a><span class=""dff dfs cfc""><b>Active Topics</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"" & strLE & _
	"<p>Active Topics are tracked by cookies. When you click on the &quot;active topics&quot; link, a page is generated listing all topics that have been posted since your last visit to these forums (or approximately 20 minutes).</p>" & strLE & _
	"</span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""edit""></a><span class=""dff dfs cfc""><b>Editing Your Posts</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>" & strLE & _
	" You may edit or delete your own posts at any time. Just go to the topic where the post to be edited or deleted is located and you will see an edit or delete icon (" & getCurrentIcon(strIconEditTopic,"Edit","class=""vam""") & getCurrentIcon(strIconDeleteReply,"Delete","class=""vam""") & ") on the line that begins &quot;posted on...&quot; Click on this icon to edit or delete the post. No one else can edit your post, except for the forum Moderator or the forum Administrator. "
if (strEditedByDate = "1") then Response.Write "A note is generated at the bottom of each edited post displaying when and by whom the post was edited."
Response.Write "</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""attach""></a><span class=""dff dfs cfc""><b>Attaching Files</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>" & strLE & _
	" For security reasons, you may not attach files to any posts. However, you may cut and paste text into your post.</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""search""></a><span class=""dff dfs cfc""><b>Searching For Specific Posts</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>" & strLE & _
	" You may search for specific posts based on a word or words found in the posts, user name, date, and particular forum(s). Simply click on the &quot;search&quot; link at the top of most pages.</span></p></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""editprofile""></a><span class=""dff dfs cfc""><b>Editing Your Profile</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>You may easily change any information stored in your registration profile by using the ""profile"" link located near the top of each page. Simply identify yourself by typing your Username and Password and all of your profile information will appear on screen. You may edit any information (except your Username).</p></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""signature""></a><span class=""dff dfs cfc""><b>Signatures</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>You may attach signatures to the end of your posts when you post either a New Topic or Reply. Your signature is editable by clicking on &quot;profile&quot; at the top of any forum page and entering your Username and Password.</p>" & strLE & _
	"<p>NOTE: HTML can't be used in Signatures.</p></span></td>" & strLE & _
	"</tr>" & strLE
if (strBadWordFilter = "1") then
	Response.Write "<tr>" & strLE & _
		"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""censor""></a><span class=""dff dfs cfc""><b>Censoring Posts</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		" The Forum does censor certain words that may be posted; however, this censoring is not an exact science, and is being done based on the words that are being screened, so certain words may be censored out of context. By default, words that are censored are replaced with asterisks.</span></p></td>" & strLE & _
	"</tr>" & strLE
end if
if (stremail = "1") then
	Response.Write "<tr>" & strLE & _
		"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""pw""></a><span class=""dff dfs cfc""><b>Lost Password</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		" Changing a lost password is simple, assuming that e-mail features are turned on for this forum. All of the pages that require you to identify yourself with your Username and Password carry a &quot;lost Password&quot; link that you can use to have a code e-mailed instantly to your e-mail address of record that will allow you to create a new password. Because of the Encryption that we use for your password, we cannot tell you what your password is.</span></p></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""notify""></a><span class=""dff dfs cfc""><b>Can I be notified by e-mail when there are new posts?</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		" Yes, the <b>Subscription</b> feature allows you to subscribe to the entire Board, individual Categories, Forums and/or Topics, depending on what the administrator of this site allows. You will receive an e-mail notifying you of a post that has been made to the Category/Forum/Topic that you have subscribed to. There are four levels of subscription:<br>" & strLE & _
		"<blockquote>" & strLE & _
		" - <b>Board Wide Subscription</b><br>" & strLE & _
		" If you can subscribe to an entire Board, you'll get a notification for any posts made within all the forums inside that board.<br><br>" & strLE & _
		" - <b>Category Wide Subscription</b><br>" & strLE & _
		" You can subscribe to an entire Category, which will notify you if there was any posts made within any topic, within any forum, within that Category.<br><br>" & strLE & _
		" - <b>Forum Wide Subscription</b><br>" & strLE & _
		" If you don't want to subscribe to an entire Category, you can subscribe to a single forum. This will notify you of any posts made within any topic, within that forum.<br><br>" & strLE & _
		" - <b>Topic Wide Subscription</b><br>" & strLE & _
		" More conveniently, you can subscribe to just an individual topic. You will be notified of any post made within that topic." & strLE & _
		"</blockquote>" & strLE & _
		" Each level of subscription is optional. The administrator can turn <b>On/Off</b> each level of subscription for each Category/Forum/Topic.<br>" & strLE & _
		" To Subscribe or Unsubscribe from any level of subscription, you can use the ""My Subscriptions"" link, located near the top of each page to manage your subscriptions. Or you can click on the subscribe/unsubscribe icons (" & getCurrentIcon(strIconSubscribe,"Subscribe","class=""vam""") & "&nbsp;" & getCurrentIcon(strIconUnsubscribe,"UnSubscribe","class=""vam""") & ") for that Category/Forum/Topic you want to subscribe/unsubscribe to/from.<br><br></span></p></td>" & strLE & _
		"</tr>" & strLE
end if
if (strModeration = "1") then
	Response.Write "<tr>" & strLE & _
		"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""moderation""></a><span class=""dff dfs cfc""><b>What does it mean if a forum has Moderation enabled?</b></span></td>" & strLE & _
		"</tr>" & strLE & _
		"<tr>" & strLE & _
		"<td>" & strLE & _
		"<p>" & strLE & _
		"<b>Moderation:</b> This feature allows the Administrator or the Moderator to ""<b>Approve</b>"", ""<b>Hold</b>"" or ""<b>Delete</b>"" a users post before it is shown to the public.<br>" & strLE & _
		"<b>Approve:</b> Only the administrators or the moderators will be able to approve a post made to a moderated forum. When the post is approved, it will be made viewable to the public.<br>" & strLE & _
		"<b>Hold:</b> When a user posts a message to a moderated forum, the message is automatically put on hold until a moderator or an administrator approves of the post. No one will be able to view the post while it is put on hold.<br>" & strLE & _
		"<i>NOTE: Authors of the post will be able to edit their post during this mode.</i><br>" & strLE & _
		"<b>Delete:</b> If the administrator or moderator chooses this option, the post will be deleted and an e-mail will be sent to the poster of the message, informing them that their post was not approved. The administrator/moderator will be able to give their reason for not approving the post in the e-mail.<br></span></p></td>" & strLE & _
		"</tr>" & strLE
end if
Response.Write "<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""COPPA""></a><span class=""dff dfs cfc""><b>What is COPPA?</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>The Children's Online Privacy Protection Act and Rule apply to individually identifiable information about a child that is collected online, such as full name, home address, e-mail address, telephone number or any other information that would allow someone to identify or contact the child. The Act and Rule also cover other types of information -- for example, hobbies, interests and information collected through cookies or other types of tracking mechanisms -- when they are tied to individually identifiable information. More information can be found <span class=""smt""><a href=""http://www.ftc.gov/bcp/conline/pubs/buspubs/coppa.htm"" title=""What is COPPA?"">here</a></span>.</p></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><a href=""#top"">" & getCurrentIcon(strIconGoUp,"Go To Top Of Page","class=""vam fr""") & "</a><a name=""GetForum""></a><span class=""dff dfs cfc""><b>Getting Your Own Forum</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<p>The most recent version of this Snitz Forum can be downloaded at <span class=""smt""><a href=""http://forum.snitz.com/"" target=""_blank"" title=""Link to Snitz Forums 2000 Homepage!"">this Internet web site</a></span>.</p>" & strLE & _
	"<p>NOTE: The software is highly configurable, and the baseline Snitz Forum may not have all the features this forum does.</p></span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE
WriteFooter
Response.End
%>
