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
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<table class=""tc"" width=""100%"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""nw l"" width=""33%""><span class=""dff dfs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_forums.asp"">Forum&nbsp;Deletion/Archival</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;Compact&nbsp;Database<br></span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE
strForumDB     = getForumDB()
strForumDB     = replace(strForumDB,";","",1,1)
strDBPath      = Left(strForumDB,InStrRev(strForumDB,"\"))
strTempFile    = strDBPath & "Snitz_compacted.mdb"
DBFolderExists = CheckDBFolder(strDBPath)
if Application(strCookieURL & "down") then status = "Closed" else status = "Open"
Response.Write "<table width=""70%"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
	"<tr>" & strLE & _
	"<td>" & strLE & _
	"<table class=""tbc"" width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
	"<tr>" & strLE & _
	"<td class=""ccc""><span class=""dff dfs cfc""><b>Administrative Forum Archive Functions - Compact DB</b></span></td>" & strLE & _
	"</tr>" & strLE & _
	"<tr>" & strLE & _
	"<td class=""fcc vat""><span class=""dff dfs dfc"">" & strLE
if request.querystring("action") = ""  then
	Response.Write "<p class=""c""><b>Depending on security settings at your Host, these operations may or may not be successful.  However, no harm should befall your data</b></p>" & strLE & _
		"Your original database will be copied to:<br><br><b>" & left(strForumDB,len(strForumDB)-4) & "_" & DateToStr(strForumTimeadjust) & ".bak" & "</b><br><br>as a backup and then compacted to:<br><br><b>" & strTempFile & "</b><br>" & strLE & _
		"<br>" & strLE & _
		"If these steps are successful, the original DB will be replaced by the compacted DB.<br>" & strLE & _
		"<br>" & strLE & _
		"This may take some time depending on the size of your database.<br>" & strLE & _
		"<br>" & strLE & _
		"<span class=""hlfc""><p class=""c"">You will have to <b>CLOSE</b> the forum while the database is being compacted.</span><br><br>" & strLE & _
		"Current Status of Forum:<br><b>" & status & "</b><br><br>" & strLE
	if Application(strCookieURL & "down") then
		Response.Write "Are you sure you want to compact the database?<br>" & strLE & _
			"<span class=""smt""><a href=""admin_compactdb.asp?action=Yes"">Yes</a></span>&nbsp;<span class=""smt""><a href=""admin_compactdb.asp?action=No"">No</a></span></p>" & strLE
	else
		Response.Write "<span class=""smt""><a href=""down.asp"">Click here</a></span> to close the forum before you start.</p>" & strLE
	end if
elseif request.querystring("action") = "No" then
	Response.Write "<p class=""c"">You have chosen not to compact your database. You can compact your database at a later time.<br><br>" & strLE & _
		"You will need to open your forums before you continue.<br>" & strLE & _
		"<span class=""smt""><a href=""down.asp"">Click here</a></span> to open your forum.</p>" & strLE
elseif request.querystring("action") = "Yes" then
	my_conn.close
	strTempConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTempFile
	if DBFolderExists = false then
		Response.Write "<br><p class=""c"">Unable to create folder:<br>" & strDBPath & ".</p><br>" & strLE
	else
		if BackupDB(strForumDB) then
			set jro = server.createobject("jro.JetEngine")
			jro.CompactDatabase strConnString, strTempConnString
			if err <> 0 then
				bError = True
				Response.Write "Error Compacting:<br>" & err.description & strLE
			else
				Response.Write "<br><p class=""c"">Database Compacted successfully.</p>" & strLE
			end if
			if not bError then
				if not RenameFile( strTempFile, strForumDB) then
					Response.Write "Error Replacing:<br>" & err.description & strLE
				else
					Response.Write "<br><p class=""c"">Database renamed successfully.</p><br>" & strLE
				end if
			end if
		else
			Response.Write "<br><p class=""c"">Unable to back up database</p><br>" & strLE
		end if
	end if
	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString
	Response.Write "<p class=""c""><span class=""smt""><a href=""down.asp"">Re-open Forum</a></span></p><br>" & strLE
end if
Response.Write "</span></td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"</td>" & strLE & _
	"</tr>" & strLE & _
	"</table>" & strLE & _
	"<br>" & strLE
'if Application(strCookieURL & "down") then
	'Response.Write "<p class=""c""><a href=""down.asp?target=admin_forums.asp""><span class=""dff dfs dfc"">Open Forum</span></a></p>" & strLE
	Response.Write "<p class=""c""><span class=""dff dfs dfc""><a href=""admin_forums.asp"">Back to Forums Administration</a></span></p>" & strLE
'end if
Response.Write "<br>" & strLE & _
		"<br>" & strLE
Call WriteFooter
Response.End
Function RenameFile(sFrom, sTo)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		RenameFile = False
		Exit Function
	end if
	fso.DeleteFile sTo,true
	fso.MoveFile sFrom, sTo
	set fso = nothing
	RenameFile = True
end Function
Function BackupDB(sFrom)
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		BackupDB = False
		Exit Function
	end if
	fso.CopyFile sFrom, fso.GetParentFolderName(sFrom) & "\" & fso.GetBaseName(sFrom) & "_" & DateTostr(strForumTimeAdjust) & ".bak", true
	set fso = nothing
	BackupDB = True
end Function
Function GetForumDB()
	dim tmpFileName
 	tmpFileName = split(strConnstring,"Source=",2,1)
	GetForumDB = tmpFileName(1)
end Function
Function CheckDBFolder(strPath)
	Dim fso, blnExists
	Set fso = CreateObject("Scripting.FileSystemObject")
	if err.number <> 0 then
		CheckDBFolder = False
		Exit Function
	end if
	blnExists = fso.FolderExists(strPath)
	if blnExists = false then
		fso.CreateFolder(strPath)
		CheckDBFolder = True
	else
		CheckDBFolder = True
	end if
End Function
%>
