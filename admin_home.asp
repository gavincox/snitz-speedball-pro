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
<!--#INCLUDE FILE="config.asp"-->
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp" -->
<!--#INCLUDE FILE="cb/admin_home_cb.asp" -->
<%
select case strDBType
	case "access"
		if instr(lcase(strConnString), lcase(Server.MapPath("snitz_forums_2000.mdb")))> 0 then
			Response.Write "<br>" & strLE & _
				"<table border=""1"" width=""100%"" bgcolor=""red"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""c""><span color=""white"" size=""2"">" & _
				"<b>WARNING:</b> The location of your access database may not be secure.<br><br>" & _
				"You should consider moving the database from <b>" & Server.MapPath("snitz_forums_2000.mdb") & "</b> to a directory not directly accessible via a URL and/or renaming the database to another name." & _
				"<br><br><i>(After moving or renaming your database, remember to change the strConnString setting in config.asp.)</i>" & _
				"</span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table><br>" & strLE
		end if
	case "sqlserver"
		if instr(lcase(strConnString), ";uid=sa;")> 0 then
			Response.Write "<br>" & strLE & _
				"<table border=""1"" width=""100%"" bgcolor=""red"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""c""><span color=""white"" size=""2"">" & _
				"<b>WARNING:</b> You are connecting to your MS SQL Server database with the <b>sa</b> user.<br><br>" & _
				"After you have completed your installation, consider creating a new user with lower privileges and use that to connect to the database instead." & _
				"</span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table><br>" & strLE
		end if
	case "mysql"
		if instr(lcase(strConnString), ";uid=root;")> 0 then
			Response.Write "<br>" & strLE & _
				"<table border=""1"" width=""100%"" bgcolor=""red"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""c""><span color=""white"" size=""2"">" & _
				"<b>WARNING:</b> You are connecting to your MySQL Server database with the <b>root</b> user.<br><br>" & _
				"After you have completed your installation, consider creating a new user with lower privileges and use that to connect to the database instead." & _
				"</span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table><br>" & strLE
		end if
end select
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Admin&nbsp;Section" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE & _
	"<table class=""admin"">" & strLE & _
	"<tr>" & strLE & _
	"<th colspan=""2""><b>Administrative Functions</b></th>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<th><b>Forum Feature Configuration</b></th>" & strLE & _
	"<th><b>Other Configuration Options and Features</b></th>" & strLE & _
	"</tr>" & strLE & _
	"<tr class=""vat"">" & strLE & _
	"<td>" & strLE & _
	"<ul class=""smt"">" & strLE & _
	"<li><a href=""admin_config_system.asp"">Main Forum Configuration</a></li>" & strLE & _
	"<li><a href=""admin_config_features.asp"">Feature Configuration</a></li>" & strLE
	if strAuthType = "nt" then Response.Write "<li><a href=""admin_config_NT_features.asp"">Feature NT Configuration</a></li>" & strLE
	Response.Write "<li><a href=""admin_config_members.asp"">Member Details Configuration</a></li>" & strLE & _
	"<li><a href=""admin_config_ranks.asp"">Ranking Configuration</a></li>" & strLE & _
	"<li><a href=""admin_config_datetime.asp"">Server Date/Time Configuration</a></li>" & strLE & _
	"<li><a href=""admin_config_email.asp"">Email Server Configuration</a></li>" & strLE
if strFilterEMailAddresses = "1" Then Response.Write "<li><a href=""admin_spamserver.asp"">Blocked E-Mail Domains</a></li>" & strLE
'"<li><a href=""admin_config_colors.asp"">Font/Table Color Code Configuration</a></li>" & strLE & _
	Response.Write "<li><a href=""javascript:openWindow3('admin_config_badwords.asp')"">Bad Word Filter Configuration</a></li>" & strLE & _
	"<li><a href=""javascript:openWindow3('admin_config_namefilter.asp')"">UserName Filter Configuration</a></li>" & strLE & _
	"<li><a href=""javascript:openWindow3('admin_config_order.asp')"">Category/Forum Order Configuration</a></li>" & strLE & _
	"<li><a href=""admin_etc.asp"">Forum Cleanup Tools</a></li>" & strLE & _
	"<li><a href=""admin_config_groupcats.asp"">Group Categories Configuration</a></li>" & strLE & _
	"</ul>" & strLE & _
	"</td>" & strLE & _
	"<td>" & strLE & _
	"<ul class=""smt"">" & strLE
if strEmailVal = "1" then Response.Write "<li><a href=""admin_accounts_pending.asp"">Members Pending</a>&nbsp;<span class=""ffs"">(" & User_Count & ")</span></li>" & strLE
Response.Write "<li><a href=""admin_members.asp"">Admin/Moderator List</a></li>" & strLE & _
	"<li><a href=""admin_member_search.asp"">Member Search</a></li>" & strLE & _
	"<li><a href=""admin_moderators.asp"">Moderator Setup</a></li>" & strLE & _
	"<li><a href=""admin_emaillist.asp"">E-mail List</a></li>" & strLE & _
	"<li><a href=""admin_info.asp"">Server Information</a></li>" & strLE & _
	"<li><a href=""admin_variable_info.asp"">Forum Variables Information</a></li>" & strLE & _
	"<li><a href=""admin_count.asp"">Update Forum Counts</a></li>" & strLE
if strArchiveState = "1" then Response.Write "<li><a href=""admin_forums.asp"">Archive Forum Topics</a></li>" & strLE
Response.Write "<li><a href=""down.asp"">Shut Down the Forum</a></li>" & strLE & _
	"<li><a href=""admin_mod_dbsetup.asp"">MOD Setup</a><span class=""ffs"">&nbsp;(<span class=""smt""><a href=""admin_mod_dbsetup2.asp"">Alternative MOD Setup</a></span>)</span></li>" & strLE & _
	"<li><a href=""setup.asp"">Check Installation</a><span class=""ffs""><b> (Run after each upgrade !)</b></span></li>" & strLE & _
	"</ul>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE
Call WriteFooter
Response.End
%>
