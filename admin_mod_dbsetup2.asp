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
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
%>
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header.asp"-->
<%
Response.Write "<table width=""100%"">" & strLE & _
		"<tr>" & strLE & _
		"<td class=""nw l"" width=""33%""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
		getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
		getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpen,"","") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
		getCurrentIcon(strIconBlank,"","") & getCurrentIcon(strIconBar,"","") & getCurrentIcon(strIconFolderOpenTopic,"","") & "&nbsp;MOD&nbsp;Setup<br></span></td>" & strLE & _
		"</tr>" & strLE & _
		"</table>" & strLE

if MemberID <> intAdminMemberID then
	Err_Msg = "<li>Only the Forum Admin can access this page</li>"

	Response.Write "<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""3"" color=""#FF0000"">There has been a problem!</span></p>" & strLE & _
			"<table class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td><span face=""Verdana, Arial, Helvetica"" size=""2"" color=""#FF0000""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""JavaScript:history.go(-1)"">Go Back To Admin Section</a></span></p>" & strLE
	WriteFooter
	Response.End
end if

Dim strTableName
Dim fieldArray (100)
Dim idFieldName
Dim tableExists
Dim fieldExists
Dim ErrorCount

tableExists   = -2147217900
tableNotExist = -2147217865
fieldExists   = -2147217887
ErrorCount = 0

on error resume next
Select case strDBType
	case "access"
		strUserDBType = "Microsoft Access 97/2000/2002"
	case "sqlserver"
		strUserDBType = "Microsoft SQL Server 6.x/7.x/2000"
	case "mysql"
		strUserDBType = "MySQL Server"
end Select

strRqMethod = Request.Form("method")

if strRqMethod = "Process" then
	if Request.Form("Message") = "" then
		Err_Msg = "<li>You did not enter any code to process</li>"

		Response.Write "<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""3"" color=""#FF0000"">There Was A Problem With Your Details</span></p>" & strLE & _
				"<table class=""tc"">" & strLE & _
				"<tr>" & strLE & _
				"<td><span face=""Verdana, Arial, Helvetica"" size=""2"" color=""#FF0000""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
		WriteFooter
		Response.End
	end if

	codetoprocess = split(Request.Form("Message"), chr(13) + chr(10))
	keycnt = ubound(codetoprocess)
	Response.Write "<p class=""c"">There were <b>" & keycnt & "</b> lines of code</p>"

	x = 0
	strModTitle = codetoprocess(x)
	Select case uCase(strModTitle)
		case "[CREATE]","[ALTER]","[DELETE]","[INSERT]","[UPDATE]","[DROP]"
			strModTitle = "Database Update"
		case else
	end select

	Response.Write "<table class=""tc"" cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"">" &_
			"<tr>" &_
			"<td bgcolor=""#9FAFDF"" align=""center"">" &_
			"<p>" &_
			"<span face=""Verdana, Arial, Helvetica"" size=""2"">"
	sqlVer = Request.Form("sqltype")
	response.write ("<span face=""Verdana, Arial, Helvetica"" size=""3"">")
	response.write ("<h4>" & ModName & "</h4></span>")
	Response.Write "<div align=""center""><p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""3"">" & strModTitle & "</span></p></div>" & strLE

	do while x < keycnt
		sectionName = codetoprocess(x)
		Select case uCase(sectionName)
			case "[CREATE]","[ALTER]","[DELETE]","[INSERT]","[UPDATE]","[DROP]"
				x = x + 1
				Select case uCase(sectionName)
					case "[CREATE]"
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						idFieldName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						CreateTables(rec)
					case "[ALTER]"
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						AlterTables(rec)
					case "[DELETE]"
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						DeleteValues(rec)
					case "[INSERT]"
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						InsertValues(rec)
					case "[UPDATE]"
						strTableName = uCase(codetoprocess(x))
						x = x + 1
						tempField = codetoprocess(x)
						rec = 0
						do while uCase(tempField) <> "[END]"
							fieldArray(rec) = tempField
							rec = rec+1
							x = x + 1
							tempField = codetoprocess(x)
						loop
						UpdateValues(rec)
					case "[DROP]"
						strTableName = codetoprocess(x)
						x = x + 1
						tempField = codetoprocess(x)
						DropTable()
				end select
				x = x + 1
			case else
				x = x + 1
		end select
	loop

	if ErrorCount > 0 then
		Response.write "<br>If there were errors please post a question in the MOD Implementation Forum at<br>"
		Response.write "<a href=""http://forum.snitz.com/forum/forum.asp?FORUM_ID=94&CAT_ID=10"">Snitz Forums</a>"
	else
		Response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""2""><p><b>Database setup finished</b></p>"
	end if
	Response.write "</span>" &_
		"<form action=""" & Request.ServerVariables("PATH_INFO") & """ method=""post"" name=""form2"">" &_
		"<input type=""hidden"" name=""modthod"" value="""">" &_
		"<input type=""submit"" name=""submit2"" value=""Finished""></form>" &_
		"</span></p></td>" &_
		"</tr>" &_
		"<tr>" &_
		"<td class=""c"">" &_
		"<span face=""Verdana, Arial, Helvetica"" size=""2"">" &_
		"<a href=""default.asp"" target=""_top"">Click here to go to the forum.</a>" &_
		"</span></td>" &_
		"</tr>" &_
		"</table>"
else
	Response.write 	"<form action=""" & Request.ServerVariables("PATH_INFO") & """ method=""post"" name=""form1"">" & strLE & _
			"<input name=""method"" type=""hidden"" value=""Process"">" & strLE & _
			"<table class=""tc"" cellspacing=""0"" cellpadding=""0"" width=""50%"" height=""50%"">" & strLE & _
			"<tr>" & strLE & _
			"<td bgcolor=""#000000"">" & strLE & _
			"<table width=""100%"" cellspacing=""1"" cellpadding=""4"">" & strLE & _
			"<tr>" & strLE & _
			"	   <td bgcolor=""#191970"" colspan=""2""><span face=""Verdana, Arial, Helvetica"" size=""1"" color=""#F5FFFA""><b>Snitz Forums 2000 MOD Database Setup</b></span></td>" & strLE & _
			"	 </tr>" & strLE & _
			"<tr>" & strLE
	If strDBType = "" then
		Response.Write "<td class=""c"" bgColor=""#9FAFDF"" ><span face=""Verdana, Arial, Helvetica"" size=""2""><br>" & strLE & _
				"<span face=""Verdana, Arial, Helvetica"" color=""#191970"" size=""2"">Your <b>strDBType</b> is not set, please edit your <b>config.asp</b> file<br>" &_
				"to reflect your database type<br></span>" & _
				"<br><a href=""admin_home.asp"">Back to Admin Options</a></span></td>" & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table><br>" & strLE
		WriteFooter
		Response.End
	end if
	Response.Write "<td bgColor=""#9FAFDF"" align=""right"" nowrap><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Database Type:</b></span></td>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strUserDBType & "</span></td>" & strLE & _
			"	 </tr>" & strLE
	If strDBType = "sqlserver" then
		Response.Write "<tr>" & strLE & _
				"<td bgColor=""#9FAFDF"" align=""right"" nowrap><span face=""Verdana, Arial, Helvetica"" size=""2""><b>SQL Server Version:</b></span></td>" & strLE & _
				"<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""1"">" & strLE & _
				"	   <input type=""radio"" name=""sqltype"" value=""6""> SQL 6.x&nbsp;&nbsp;&nbsp;" & strLE & _
				"	   <input type=""radio"" name=""sqltype"" value=""7"" checked> SQL 7.x/2000</span></td>" & strLE & _
				"</tr>" & strLE
	end if
	Response.Write "<tr>" & strLE & _
			"<td class=""r"" bgColor=""#9FAFDF"" valign=""top""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>CODE:</b></span></td>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""center""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
			"<textarea cols=""50"" name=""Message"" rows=""10"" wrap=""VIRTUAL""></textarea></span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""center"" colspan=""2""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
			"	   <span face=""Verdana, Arial, Helvetica"" size=""1"">Enter the Code in the box above that you would like to process.<br>A script will execute to perform the database upgrade.</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""center"" colspan=""2""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
			"<input name=""Submit"" type=""submit"" value=""Submit"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset Fields""></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE
end if
WriteFooter
Response.End

Sub CreateTables( numfields )
	response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	response.write "<b>Creating table(s)...</b><br>"
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		TablePrefix = strMemberTablePrefix
	else
		TablePrefix = strTablePrefix
	end if

	strSql = "CREATE TABLE " & TablePrefix & strTableName & "( "
	if idFieldName <> "" then
		select case strDBType
			case "access"
				if Instr(strConnString,"(*.mdb)") then
					strSql = strSql & idFieldName &" COUNTER CONSTRAINT PrimaryKey PRIMARY KEY "
				else
					strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
				end if
			case "sqlserver"
				strSql = strSql & idFieldName &" int IDENTITY (1, 1) PRIMARY KEY NOT NULL "
			case "mysql"
				strSql = strSql & idFieldName &" INT (11) NOT NULL auto_increment "
		end select
	end if
	for y = 0 to numfields -1
	on error resume next
		tmpArray = split(fieldArray(y),"#")
		fName = uCase(tmpArray(0))
		fType = lCase(tmpArray(1))
		fNull = uCase(tmpArray(2))
		fDefault = tmpArray(3)
		if idFieldName <> "" or y <> 0 then
			strSql = strSql & ", "
		end if
		select case strDBType
			case "access"
				fType = replace(fType,"varchar (","text (")
			case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
			case "mysql"
				fType = replace(fType,"memo","text")
				fType = replace(fType,"#int","#int (11)")
				fType = replace(fType,"#smallint","#smallint (6)")
		end select
		if fNull <> "NULL" then fNull = "NOT NULL"
		strSql = strSql & fName & " " & fType & " " & fNull & " "
		if fdefault <> "" then
			select case strDBType
				case "access"
					if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
				case else
					strSql = strSql & "DEFAULT " & fDefault
			end select
		end if
	next
	if strDBType = "mysql" then
		if idFieldName <> "" then
			strSql = strSql & ",KEY " & TablePrefix & strTableName & "_" & idFieldName & "(" & idFieldName & "))"
		else
			strSql = strSql & ")"
		end if
	else
		strSql = strSql & ")"
	end if
	response.write strSql & "<br>"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	if err.number <> 0 and err.number <> 13 and err.number <> tableExists then
		response.write strSql & "<br>"
		response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableExists then
			response.write("<span color=""#FF0000""><b>Table already exists</b></span><br>")
		else
			response.write("<b>Table created successfully</b><br>")
		end if
	end if

	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

Sub AlterTables(numfields)
	Response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	for y = 0 to numfields -1
		on error resume next
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			TablePrefix = strMemberTablePrefix
		else
			TablePrefix = strTablePrefix
		end if
		strSql = "ALTER TABLE " & TablePrefix & strTableName
		tmpArray = split(fieldArray(y),"#")
		fAction = uCase(tmpArray(0))
		fName = uCase(tmpArray(1))
		fType = lCase(tmpArray(2))
		fNull = uCase(tmpArray(3))
		fDefault = tmpArray(4)
		select case fAction
			case "ADD"
				strSQL = strSQL & " ADD "
				if strDBType = "access" then strSql = strSql & "COLUMN "
			case "DROP"
				strSQL = strSQL & " DROP COLUMN "
			case "ALTER"
				strSQL = strSQL & " ALTER COLUMN "
			case else
		end select
		if fAction = "ADD" or fAction = "ALTER" then
			select case strDBType
				case "access"
					fType = replace(fType,"varchar (","text (")
				case "sqlserver"
				select case sqlVer
					case 7
						fType = replace(fType,"memo","ntext")
						fType = replace(fType,"varchar","nvarchar")
						fType = replace(fType,"date","datetime")
					case else
						fType = replace(fType,"memo","text")
				end select
				case "mysql"
					fType = replace(fType,"memo","text")
					fType = replace(fType,"#int","#int (11)")
					fType = replace(fType,"#smallint","#smallint (6)")
			end select
			if fNull <> "NULL" then fNull = "NOT NULL"
			strSql = strSQL & fName & " " & fType & " " & fNULL & " "
			if fDefault <> "" then
				select case strDBType
					case "access"
						if Instr(lcase(strConnString), "jet") then strSql = strSql & "DEFAULT " & fDefault
					case else
						strSql = strSql & "DEFAULT " & fDefault
				end select
			end if
			response.write "<b>" & LCase(fAction) & "ing Column " & fName & "...</b><br>"
		else
			strSql = strSQL & fName
			response.write "<b>Dropping Column...</b><br>"
		end if
		response.write strSql & "<br>"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		if err.number <> 0 and err.number <> 13 and err.number <> fieldExists then
			response.write strSQL & "<br>"
			response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
			ErrorCount = ErrorCount + 1
			resultString = ""
		else
			if fAction = "DROP" then
				response.write("<b>Column " & LCase(fAction) & "ped successfully</b><br>")
				resultString = "<b>Table(s) updated</b><br>"
			else
				if err.number = fieldExists then
					response.write("<b><span color=""#FF0000"">Column already exists</span></b><br>")
					resultString = ""
				else
					response.write("<b>Column " & LCase(fAction) & "ed successfully</b><br>")
				end if
			end if
			if fDefault <> "" and err.number <> fieldExists then
				strSQL = "UPDATE " & TablePrefix & strTableName & " SET " & fName & "=" & fDefault
				response.write strSql & "<br>"
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				response.write "<b>Populating Current Records with new Default value</b><br>"
				resultString = "<b>Table(s) updated</b><br>"
			end if
		end if

		if fieldArray(y) = "" then y = numfields
	next
	response.write(resultString)
	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

Sub InsertValues(numfields)
	Response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	on error resume next
	response.write ("<b>Adding new records..</b><br>")
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "INSERT INTO " & strMemberTablePrefix & strTableName & " "
		else
			strSql = "INSERT INTO " & strTablePrefix & strTableName & " "
		end if
		tmpArray = split(fieldArray(y),"#")
		fNames = tmpArray(0)
		fValues = tmpArray(1)
		strSql = strSql & tmpArray(0) & " VALUES " & tmpArray(1)
		response.write strSql & "<br>"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	next

	if err.number <> 0 and err.number <> 13 then
		response.write strSql & "<br>"
		response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<br><b>Value(s) updated successfully</b>")
	end if
	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

Sub UpdateValues(numfields)
	on error resume next
	Response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	response.write ("<b>Updating Forum Values..</b><br>")
	for y = 0 to numfields-1
		if Instr(1,strTableName,"MEMBER",1) > 0 then
			strSql = "UPDATE " & strMemberTablePrefix & strTableName & " SET"
		else
			strSql = "UPDATE " & strTablePrefix & strTableName & " SET"
		end if
		tmpArray = split(fieldArray(y),"#")
		fName = tmpArray(0)
		fValue = tmpArray(1)
		fWhere = tmpArray(2)
		strSql = strSql & " " & fName & " = " & fvalue
		if fWhere <> "" then
			strSql = strSql & " WHERE " & fWhere
		end if
		response.write strSql & "<br>"
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
	next

	if err.number <> 0 then
		response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
		ErrorCount = ErrorCount + 1
		response.write strSql & "<br>"
	else
		response.write("<br><b>Value(s) updated successfully</b>")
	end if
	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

Sub DeleteValues(numfields)
	on error resume next
	response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	response.write ("<b>Updating Forum Values..</b><br>")
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DELETE FROM " & strMemberTablePrefix & strTableName & " WHERE "
	else
		strSql = "DELETE FROM " & strTablePrefix & strTableName & " WHERE "
	end if
	tmpArray = fieldArray(0)
	strSql = strSql & tmpArray
	response.write strSql & "<br>"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 then
		response.write strSql & "<br>"
		response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
		ErrorCount = ErrorCount + 1
	else
		response.write("<br><b>Value(s) updated successfully</b>")
	end if
	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

Sub DropTable()
	on error resume next
	response.write "<br><span face=""Verdana, Arial, Helvetica"" size=""1"">"
	response.write ("<b>Dropping Table..</b><br>")
	if Instr(1,strTableName,"MEMBER",1) > 0 then
		strSql = "DROP TABLE " & strMemberTablePrefix & strTableName
	else
		strSql = "DROP TABLE " & strTablePrefix & strTableName
	end if
	response.write strSql & "<br>"
	my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

	if err.number <> 0 and err.number <> 13 and err.number <> tableNotExist then
		response.write strSql & "<br>"
		response.write("<span color=""#FF0000"">" & err.number & " | " & err.description & "</span><br>")
		ErrorCount = ErrorCount + 1
	else
		if err.number = tableNotExist then
			response.write("<br><b>Table does not exist</b>")
		else
			response.write("<br><b>Table dropped successfully</b>")
		end if
	end if
	response.write("<hr size=""1"" width=""260"" align=""center"" color=""blue""></span>")
end Sub

on error goto 0
%>
