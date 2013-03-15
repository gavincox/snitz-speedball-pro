& strTablePrefix & "TOPICS MODIFY T_REPLIES DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			strOkMessage = "T_REPLIES field in " & strTablePrefix & "A_TOPICS table has been changed"

			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "A_TOPICS MODIFY T_REPLIES DOUBLE "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			strOkMessage = "F_TOPICS field in " & strTablePrefix & "FORUM table has been changed"

			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_TOPICS DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			strOkMessage = "F_A_TOPICS field in " & strTablePrefix & "FORUM table has been changed"

			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_A_TOPICS DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			strOkMessage = "F_COUNT field in " & strTablePrefix & "FORUM table has been changed"

			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_COUNT DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			strOkMessage = "F_A_COUNT field in " & strTablePrefix & "FORUM table has been changed"

			SpecialSQL12(MySql) = "ALTER TABLE " & strTablePrefix & "FORUM MODIFY F_A_COUNT DOUBLE DEFAULT '0' "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush

			'## Add the M_ALLOWEMAIL field to MEMBERS and MEMBERS_PENDING
			strOkMessage = "M_ALLOWEMAIL has been added to " & strMemberTablePrefix & "MEMBERS"

			SpecialSQL12(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL6)   = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL7)   = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(MySql)  = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS ADD M_ALLOWEMAIL smallint DEFAULT 0"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			'Update defaults?

			strOkMessage = "M_ALLOWEMAIL has been added to " & strMemberTablePrefix & "MEMBERS_PENDING"

			SpecialSQL12(Access) = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL6)   = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(SQL7)   = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0 "
			SpecialSQL12(MySql)  = "ALTER TABLE " & strMemberTablePrefix & "MEMBERS_PENDING ADD M_ALLOWEMAIL smallint DEFAULT 0"
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush
			'Update defaults?

			'## Add Table SPAM_MAIL to the database
	 		strOkMessage = "Table " & strFilterTablePrefix & "SPAM_MAIL created"

			SpecialSQL12(Access) = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(Access) = SpecialSQL12(Access) & "SPAM_ID COUNTER CONSTRAINT PrimaryKey PRIMARY KEY , "
			SpecialSQL12(Access) = SpecialSQL12(Access) & "SPAM_SERVER text (255) NULL )"

			SpecialSQL12(SQL6)   = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(SQL6)   = SpecialSQL12(SQL6) & "SPAM_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSQL12(SQL6)   = SpecialSQL12(SQL6) & "SPAM_SERVER varchar (255) NULL )"

			SpecialSQL12(SQL7)   = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(SQL7)   = SpecialSQL12(SQL7) & "SPAM_ID int IDENTITY (1, 1) PRIMARY KEY NOT NULL , "
			SpecialSQL12(SQL7)   = SpecialSQL12(SQL7) & "SPAM_SERVER nvarchar (255) NULL )"

			SpecialSQL12(MySql)  = "CREATE TABLE " & strFilterTablePrefix & "SPAM_MAIL ( "
			SpecialSQL12(MySql)  = SpecialSQL12(MySql) & "SPAM_ID int (11) NOT NULL auto_increment , "
			SpecialSQL12(MySql)  = SpecialSQL12(MySql) & "SPAM_SERVER VARCHAR (255) NULL , "
			SpecialSQL12(MySql)  = SpecialSQL12(MySql) & "PRIMARY KEY (SPAM_ID)) "
			call SpecialUpdates(SpecialSQL12, strOkMessage)
			Response.Flush()


			'## Add the new config-values to the database
			strDummy = SetConfigValue(0,"STRREQAIM","0")
			strDummy = SetConfigValue(0,"STRREQICQ","0")
			strDummy = SetConfigValue(0,"STRREQMSN","0")
			strDummy = SetConfigValue(0,"STRREQYAHOO","0")
			strDummy = SetConfigValue(0,"STRREQFULLNAME","0")
			strDummy = SetConfigValue(0,"STRREQPICTURE","0")
			strDummy = SetConfigValue(0,"STRREQCITY","0")
			strDummy = SetConfigValue(0,"STRREQSTATE","0")
			strDummy = SetConfigValue(0,"STRREQAGE","0")
			strDummy = SetConfigValue(0,"STRREQAGEDOB","0")
			strDummy = SetConfigValue(0,"STRREQHOMEPAGE","0")
			strDummy = SetConfigValue(0,"STRREQCOUNTRY","0")
			strDummy = SetConfigValue(0,"STRREQOCCUPATION","0")
			strDummy = SetConfigValue(0,"STRREQBIO","0")
			strDummy = SetConfigValue(0,"STRREQHOBBIES","0")
			strDummy = SetConfigValue(0,"STRREQLNEWS","0")
			strDummy = SetConfigValue(0,"STRREQQUOTE","0")
			strDummy = SetConfigValue(0,"STRREQMARSTATUS","0")
			strDummy = SetConfigValue(0,"STRREQFAVLINKS","0")
			strDummy = SetConfigValue(0,"INTMAXPOSTSTOEMAIL","10")
			strDummy = SetConfigValue(0,"STRNOMAXPOSTSTOEMAIL","You do not have enough posts to email other members. If you feel that you have received this message in error, please contact the forum administrator.")
			strDummy = SetConfigValue(0,"STRFILTEREMAILADDRESSES","0")

			'## update the version info...
			strDummy = SetConfigValue(1,"strVersion", strNewVersion) '## make sure the string is there

			strSql = "UPDATE " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " SET C_VALUE =  '" & strNewVersion & "'"
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			on error resume next
			my_Conn.Errors.Clear
			Err.Clear
			my_Conn.Execute (strSql)

			Response.Write("<table cellspacing=""0"" cellpadding=""5"" width=""50%"" class=""tc"">" & strLE)

			UpdateErrorCode = UpdateErrorCheck()

			on error goto 0

			if UpdateErrorCode = 0 then
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""green"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2""> Added default values for new fields in CONFIG table</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
			elseif UpdateErrorCode = 2 then
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">Can't add default values for new fields in CONFIG table!</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intCriticalErrors = intCriticalErrors + 1
			else
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & false & " while trying to add default values to the CONFIG table</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intCriticalErrors = intCriticalErrors + 1
			end if
			Response.Write("</table>" & strLE)
			Response.Flush
		end if

'##########################################################################################################################
'##
'## end of update section, for newer versions add below this line (and then move this notice to the end of the new section)
'##
'##########################################################################################################################

		my_Conn.Close
		set my_Conn = nothing

		if intCriticalErrors = 0 then

			Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""4"">The Upgrade has been completed without errors !</span></p>" & strLE
		else
			Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""4""><b>The Upgrade has NOT been completed without errors !</b></span></p>" & strLE & _
					"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intCriticalErrors & "  Critical Errors...</span></p>" & strLE
			if intWarnings > 0 then
				Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intWarnings & "  noncritical errors...</span></p>" & strLE
			end if
		end if
		if intCriticalErrors > 0 then
			Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp?RC=3"">Click here to retry....</a></span></p>" & strLE
		end if
		Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""default.asp"" target=""_top"">Click here to return to the forum....</a></span></p>" & strLE & _
				"</center></div>" & strLE

		Application(strCookieURL & "ConfigLoaded")= ""

	else
		Response.Redirect "setup_login.asp"
	end if

elseif ResponseCode = 5 then '## install a new database

	if strDBType = "access" then
		set mydbms_Conn = Server.CreateObject("ADODB.Connection")
		mydbms_Conn.Open strConnString
		strDBMSName     = lcase(mydbms_Conn.Properties("DBMS Name"))
		mydbms_Conn.close
		set mydbms_Conn = nothing
	end if

	Response.Write "<div align=""center""><center>" & strLE & _
			"<p><span face=""Verdana, Arial, Helvetica"" size=""4"">Installation of forum-tables in the database.</span></p>" & strLE & _
			"</center></div>" & strLE & _
			"<form action=""setup.asp?RC=6"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
			"<table cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" class=""tc"">" & strLE & _
			"<tr>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""left"">" & strLE & _
			"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">Database Type:&nbsp;&nbsp;<b>"
	if strDBType = "access" then
		Response.Write("Microsoft Access")
		select case strDBMSName
			case "access"
				Response.Write(" (using ODBC Driver)")
			case "ms jet"
				Response.Write(" (using OLEDB Driver)")
		end select
	end if
	if strDBType = "sqlserver" then Response.Write("Microsoft SQL Server")
	if strDBType = "mysql" then Response.Write("MySQL")
	Response.Write "</b></span></p></td>" & strLE & _
			"</tr>" & strLE
	if strDBType = "sqlserver" then
		Response.Write "<tr>" & strLE & _
				"<td bgColor=""#9FAFDF"" align=""left""><p>" & strLE & _
				"<p><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Select the SQL-server version you are using:</b></span></p>" & strLE & _
				"<p><span face=""Verdana, Arial, Helvetica"" size=""2""><input type=""radio"" class=""radio"" name=""SQL_Server"" value=""SQL6"" >SQL-Server 6.5<br>" & strLE & _
				"<input type=""radio"" class=""radio"" checked name=""SQL_Server"" value=""SQL7"">SQL-Server 7 / 2000&nbsp;&nbsp;&nbsp;</p></span></p></td>" & strLE & _
				"</tr>" & strLE
	end if
	if strDBType <> "access" then
		Response.Write "<tr>" & strLE & _
				"<td bgColor=""#9FAFDF"" align=""left"">" & strLE & _
				"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">To install the tables in the database you need to create the empty database on the server first.  " & strLE & _
				"    Then you have to provide a username and password of a user that has table creation/modification rights at the database you use.  " & strLE & _
				"    This might not be the same user as you use in your connectionstring !</span></p>" & strLE & _
				"<table align=""left"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
				"<tr>" & strLE & _
				"<td class=""r""><span face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Name:</b>&nbsp;</span></td>" & strLE & _
				"<td align=""left""><input type=""text"" name=""DBUserName"" size=""25"" style=""width:150px;""></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""r""><span face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password:</b>&nbsp;</span></td>" & strLE & _
				"<td align=""left""><input type=""password"" name=""DBPassword"" size=""25"" style=""width:150px;""></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE
	end if
	Response.Write "<tr>" & strLE & _
			"<td bgColor=""#9FAFDF"" align=""left"">" & strLE & _
			"<p><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Forum Admin UserName/Password:</b></span></p>" & strLE & _
			"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">Here you will choose the Forum Admin UserName & Password that will be entered into the database for the Forum Admin.  " & strLE & _
			"    The password should be something that you can remember, but not something easily guessed by anyone else.  Size limit is 25 characters.</span></p>" & strLE & _
			"<table align=""left"" cellspacing=""0"" cellpadding=""0"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""r""><span face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Username:</b>&nbsp;</span></td>" & strLE & _
			"<td align=""left""><input maxLength=""25"" type=""text"" name=""AdminName"" size=""25"" style=""width:150px;""></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""r""><span face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password:</b>&nbsp;</span></td>" & strLE & _
			"<td align=""left""><input maxLength=""25"" type=""password"" name=""AdminPassword"" size=""25"" style=""width:150px;""></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""r""><span face=""Verdana, Arial, Helvetica"" size=""2"">&nbsp;<b>Password Again:</b>&nbsp;</span></td>" & strLE & _
			"<td align=""left""><input maxLength=""25"" type=""password"" name=""AdminPassword2"" size=""25"" style=""width:150px;""></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""c""><input type=""submit"" value=""Continue"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE

elseif ResponseCode = 6 then '## start installing the tables in the database

	Err_Msg = ""

	strAdminName      = trim(chkString(Request.Form("AdminName"),"SQLString"))
	strAdminPassword  = trim(chkString(Request.Form("AdminPassword"),"SQLString"))
	strAdminPassword2 = trim(chkString(Request.Form("AdminPassword2"),"SQLString"))

	if strAdminName = "" then
		Err_Msg = Err_Msg & "<li>You must choose the Forum Admin UserName</li>"
	end if

	if Len(strAdminName) < 3 then
		Err_Msg = Err_Msg & "<li>Your Forum Admin UserName must be at least <strong>3</strong> characters long</li>"
	end if

	if not IsValidString(strAdminName) then
		Err_Msg = Err_Msg & "<li>You may not use any of these chars in the Forum Admin UserName  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
	end if

	if strAdminPassword = "" then
		Err_Msg = Err_Msg &  "<li>You must choose the Forum Admin Password</li>"
	end if

	if strAdminPassword <> strAdminPassword2 then
		Err_Msg = Err_Msg & "<li>Your Forum Admin Passwords didn't match</li>"
	end if

	if not IsValidString(strAdminPassword) then
		Err_Msg = Err_Msg & "<li>You may not use any of these chars in the Forum Admin Password  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
	end if

	if Err_Msg <> "" then
		Response.Write "<table height=""50%"" class=""tc"">" & strLE & _
				"<tr>" & strLE & _
				"<td>" & strLE & _
				"<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""4"" color=""#FF0000"">There has been a problem!</span></p>" & strLE & _
				"<table class=""tc"">" & strLE & _
				"<tr>" & strLE & _
				"<td><span face=""Verdana, Arial, Helvetica"" size=""2"" color=""#FF0000""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"<p class=""c""><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""JavaScript:history.go(-1)"">Go back to correct the problem</a></span></p>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE
		Response.End
	end if

	strAdminPassword = sha256("" & strAdminPassword)

	on error resume next

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strConnString

	if strDBType = "access" then
		strDBMSName = lcase(my_Conn.Properties("DBMS Name"))
	end if

	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc   = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then
			my_Conn.Errors.Clear
			Err.Clear
			Response.Redirect "setup.asp?RC=1&CC=1&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc)
		end if
	next

	my_Conn.Errors.Clear
	Err.Clear

	'## Forum_SQL
	strSql = "SELECT MEMBER_ID "
	strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS "
	strSql = strSql & " WHERE M_LEVEL = 3"

	Set rs = my_Conn.Execute(strSql)

	blnError = FALSE
	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		if ConnErrorNumber <> 0 then
			my_Conn.Errors.Clear
			Err.Clear
			blnError = TRUE
		end if
	next

	If not(blnError) then
		if (not(rs.BOF or rs.EOF) and (Session(strCookieURL & "Approval") <> "15916941253") ) then
			if strDBType = "access" then
				Response.Write "<div align=""center""><center>" & strLE & _
						"<p><span face=""Verdana, Arial, Helvetica"" size=""4"">The Forum Tables have already been installed.</span></p>" & strLE & _
						"<p><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""default.asp"">Click here to go to the Forum</a></span></p>" & strLE & _
						"</center></div>" & strLE
				Response.end
			end if
			Response.Write "<div align=""center""><center>" & strLE & _
					"<p><span face=""Verdana, Arial, Helvetica"" size=""4"">You need to logon first.</span></p>" & strLE & _
					"</center></div>" & strLE & _
					"<form action=""setup_login.asp"" method=""post"" id=""Form1"" name=""Form1"">" & strLE & _
					"<input type=""hidden"" name=""setup"" value=""Y"">" & strLE & _
					"<input type=""hidden"" name=""ReturnTo"" value=""RC=5"">" & strLE & _
					"<table cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td bgColor=""#9FAFDF"" align=""left"">" & strLE & _
					"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
					"    To Re-install the tables you need to be logged on as a forum administrator.<br>" & strLE
			if strSender <> "" then
				Response.Write "    If you are not the Administrator of this forum<br> please report this error here: <a href=""mailto:" & strSender & """>" & strSender & "</a>.<br><br>" & strLE
			end if
			Response.Write "</span></p></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td>" & strLE & _
					"<table cellspacing=""2"" cellpadding=""0"" class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td class=""c"" colspan=""2"" bgColor=""#9FAFDF""><b><span face=""Verdana, Arial, Helvetica"" size=""2"">Admin Login</span></b></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td align=""right"" nowrap><b><span face=""Verdana, Arial, Helvetica"" size=""2"">UserName:</span></b></td>" & strLE & _
					"<td><input type=""text"" name=""Name""></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td align=""right"" nowrap><b><span face=""Verdana, Arial, Helvetica"" size=""2"">Password:</span></b></td>" & strLE & _
					"<td><input type=""Password"" name=""Password""></td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td colspan=""2"" align=""right""><input type=""submit"" value=""Login"" id=""Submit1"" name=""Submit1""></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"</form>" & strLE
			Response.end
		end if
	end if

	rs.close
	Set rs = nothing

	my_Conn.Errors.Clear
	Err.Clear

	on error goto 0

	Response.Write "<div align=""center""><center>" & strLE & _
			"<p><span face=""Verdana, Arial, Helvetica"" size=""4"">Please Wait until the installation has been completed !</span></p>" & strLE

	if strDBType = "access" or not Instr(strConnString,"uid=") > 0 then
		strInstallString = strConnString
	else
		strInstallString = CreateConnectionString(strConnString, Request.Form("DBUserName"), Request.Form("DBPassword"))
	end if

	on error resume next

	set my_Conn = Server.CreateObject("ADODB.Connection")
	my_Conn.Open strInstallString

	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber = Err.Number
		ConnErrorDesc   = my_conn.Errors(counter).Description
		if ConnErrorNumber <> 0 then
			my_Conn.Errors.Clear
			Err.Clear
			Response.Redirect "setup.asp?RC=1&EC=" & ConnErrorNumber & "&ED=" & Server.URLEncode(ConnErrorDesc) & "&RET=" & Server.URLEncode("setup.asp?RC=5")
		end if
	next

	on error goto 0

	intCriticalErrors = 0

	Response.Write "<table cellspacing=""0"" cellpadding=""5"" width=""50%"" height=""50%"" class=""tc"">" & strLE & _
			"<tr>" & strLE
	strSQL_Server = Request.Form("Sql_Server")
	if strDBType = "mysql" then
%>
		<!--#INCLUDE FILE="inc_create_forum_mysql.asp" -->
<%
	elseif strDBType = "sqlserver" then
		if strSQL_Server = "SQL6" then
			strN = ""
		else
			strN = "n"
		end if
%>
		<!--#INCLUDE FILE="inc_create_forum_mssql.asp" -->
<%
	elseif strDBType = "access" then
		if strDBMSName = "ms jet" then
			strN = "n"
		else
			strN = ""
		end if
%>
		<!--#INCLUDE FILE="inc_create_forum_access.asp" -->
<%
	end if
	Response.Write "</tr>" & strLE & _
			"</table>" & strLE

	my_Conn.Close
	set my_Conn = nothing

	if intCriticalErrors = 0 then
		Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""4"">The Installation has been completed !</span></p>" & strLE
	else
		Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""4""><b>The Installation has NOT been completed !</b></span></p>" & strLE & _
				"<p><span face=""Verdana, Arial, Helvetica"" size=""2"">There were " & intCriticalErrors & "  Critical Errors...</span></p>" & strLE
	end if
	if intCriticalErrors > 0 then
		Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp?RC=5"">Click here to retry....</a></span></p>" & strLE
	end if
	Response.Write "<p><span face=""Verdana, Arial, Helvetica"" size=""2""><a href=""setup.asp"" target=""_top"">Click here to check the Database....</a></span></p>" & strLE & _
			"</center></div>" & strLE
else
	Response.Write "<html>" & strLE & _
			vbNewLine & _
			"<head>" & strLE & _
			"<title>Forum-Setup Page</title>" & strLE

	'## START - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT
	Response.Write "<meta name=""copyright"" content=""This Forum code is Copyright (C) 2000-09 Michael Anderson, Pierre Gorissen, Huw Reddick and Richard Kinser, Non-Forum Related code is Copyright (C) " & strCopyright & """>" & strLE
	'## END   - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE SNITZ FORUMS 2000 LICENSE AGREEMENT

	Response.Write "<style><!--" & strLE & _
			"a:link    {color:darkblue;text-decoration:underline}" & strLE & _
			"a:visited {color:blue;text-decoration:underline}" & strLE & _
			"a:hover   {color:red;text-decoration:underline}" & strLE & _
			"--></style>" & strLE & _
			"</head>" & strLE & _
			vbNewLine & _
			"<body bgColor=""white"" text=""midnightblue"" link=""darkblue"" aLink=""red"" vLink=""red"" onLoad=""window.focus()"">" & strLE & _
			"<div align=""center""><center>" & strLE & _
			"<span face=""Verdana, Arial, Helvetica"" size=""4"">There has been an error !!</span>" & strLE & _
			"</center></div>" & strLE & _
			"<table cellspacing=""0"" cellpadding=""5"" width=""100%"" align=""left"">" & strLE & _
			"<tr>" & strLE & _
			"<td bgColor=""#9FAFDF"" class=""c"">" & strLE & _
			"<span face=""Verdana, Arial, Helvetica"" size=""2"">"
	Response.Write(HEX(Err.number) & ", " & Err.description)
	Response.Write "</span></td>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""c"">" & strLE & _
			"<span face=""Verdana, Arial, Helvetica"" size=""2"">" & strLE & _
			"<a href=""default.asp"" target=""_top"">Click here to retry.</a>" & strLE & _
			"</span></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE
end if
Response.Write "</body>" & strLE & _
		vbNewLine & _
		"</html>" & strLE

sub CheckSqlError()

	dim ChkConnErrorNumber

	for counter = 0 to my_Conn.Errors.Count -1
		ChkConnErrorNumber = Err.Number
		if ChkConnErrorNumber <> 0 then

			my_Conn.Errors.Clear
			Err.Clear

			strSql = "SELECT " & strTablePrefix & "CONFIG.C_STRVERSION, "
			strSql = strSql & strTablePrefix & "CONFIG.C_STRSENDER "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG "

			set rsInfo = my_Conn.Execute (StrSql)
			strVersion = rsInfo("C_STRVERSION")
			strSender  = rsInfo("C_STRSENDER")

			rsInfo.Close
			set rsInfo  = nothing
			my_Conn.Close
			set my_Conn = nothing

			Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & ChkConnErrorNumber
		end if
	next
end sub

sub CheckSqlErrorNew()

	dim ChkConnErrorNumber

	for counter = 0 to my_Conn.Errors.Count -1
		ChkConnErrorNumber = Err.Number
		if ChkConnErrorNumber <> 0 then

			my_Conn.Errors.Clear
			Err.Clear

			strSql = "SELECT C_VALUE "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " WHERE C_VARIABLE = 'STRVERSION'"

			set rsInfo = my_Conn.Execute (StrSql)
			strVersion = rsInfo("C_VALUE")

			strSql = "SELECT C_VALUE "
			strSql = strSql & " FROM " & strTablePrefix & "CONFIG_NEW "
			strSql = strSql & " WHERE C_VARIABLE = 'STRSENDER'"

			set rsInfo = my_Conn.Execute (StrSql)
			strSender  = rsInfo("C_VALUE")

			rsInfo.Close
			set rsInfo  = nothing
			my_Conn.Close
			set my_Conn = nothing

			Response.Redirect "setup.asp?RC=2&MAIL=" & Server.UrlEncode(strSender) & "&VER=" & Server.URLEncode(strVersion) & "&EC=" & ChkConnErrorNumber
		end if
	next
end sub

function UpdateErrorCheck()

	dim intErrorNumber
	dim counter

	intErrorNumber = 0
	for counter = 0 to my_Conn.Errors.Count -1
		intErrorNumber = my_Conn.Errors(counter).Number
		if intErrorNumber <> 0 or Err.Number <> 0 then
			select case intErrorNumber
				case -2147217900, -2147217887
					UpdateErrorCheck = 1
					counter = my_Conn.Errors.Count -1
				case -2147467259
					UpdateErrorCheck = 2
					if strDBType = "mysql" then
						if instr(my_Conn.Errors(counter).Description, "Duplicate column name") > 0 then
							UpdateErrorCheck = 1
						end if
					end if
					counter = my_Conn.Errors.Count -1
				case else
					UpdateErrorCheck = intErrorNumber
			end select
		end if
	next
end function

Sub AddColumns(Columns, intCriticalErrors, intWarnings)

	Dim colCounter

	Response.Write("<table cellspacing=""1"" cellpadding=""5"" width=""50%"" height=""50%"" class=""tc"">" & strLE)
	For colCounter = 0 to Ubound(Columns, 1)
		on error resume next
		my_Conn.Errors.Clear
		Err.Clear

		strUpdateSql = "ALTER TABLE " & Columns(colCounter, Prefix) & Columns(colCounter, TableName) & "  "
		if strDBType = "access" then
			strUpdateSql = strUpdateSql & " ADD COLUMN " & Columns(colCounter, FieldName) & " "
		else
			strUpdateSql = strUpdateSql & " ADD " & Columns(colCounter, FieldName) & " "
		end if
		if strDBType = "access" then
			strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_Access) & " " & Columns(colCounter, ConstraintAccess) & " "
		elseif strDBType = "sqlserver" then
			if strSQL_Server = "SQL7" then
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL7) & " " & Columns(colCounter, ConstraintSQL7) & " "
			else
				strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_SQL6) & " " & Columns(colCounter, ConstraintSQL6) & " "
			end if
		elseif strDBType = "mysql" then
			strUpdateSql = strUpdateSql & " " & Columns(colCounter, DataType_MySql) & " " & Columns(colCounter, ConstraintMySql) & " "
		end if

		my_Conn.Execute strUpdateSql

		UpdateErrorCode = UpdateErrorCheck()

		on error goto 0

		if UpdateErrorCode = 0 then
			Response.Write("<tr>" & strLE)
			Response.Write("<td bgColor=""green"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded:</b></span></td>" & strLE)
			Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & Columns(colCounter, FieldName) & " has been added to the " & Columns(colCounter, TableName) & " table</span></td>" & strLE)
			Response.Write("</tr>" & strLE)
		elseif UpdateErrorCode = 1 then
			Response.Write("<tr>" & strLE)
			Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Noncritical error: </b></span></td>" & strLE)
			Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & Columns(colCounter, Fieldname) & " already existed in the " & Columns(colCounter, TableName) & " table</span></td>" & strLE)
			Response.Write("</tr>" & strLE)
			intWarnings = intWarnings + 1
		elseif UpdateErrorCode = 2 then
			Response.Write("<tr>" & strLE)
			Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: </b></span></td>" & strLE)
			Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2""> No write access to the table " & Columns(colCounter, TableName) & " <br>" & Columns(colCounter, Fieldname) & " not added to database!</span></td>" & strLE)
			Response.Write("</tr>" & strLE)
			intCriticalErrors = intCriticalErrors + 1
		else
			Response.Write("<tr>" & strLE)
			Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></span></td>" & strLE)
			Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " in statement [" & strUpdateSql & "] while trying to add " & Columns(colCounter, Fieldname) & " to the " & Columns(colCounter, TableName) & " table</span></td>" & strLE)
			Response.Write("</tr>" & strLE)
			intCriticalErrors = intCriticalErrors + 1
		end if
		Response.Flush
	Next
	Response.Write("</table>" & strLE)
end sub

sub SpecialUpdates(strSql, strOkMessage)

	dim strUpdateSql
	dim SpecialErrors

	on error resume next
	my_Conn.Errors.Clear
	Err.Clear

	if strDBType = "access" then
		strUpdateSql = strSql(Access)
	elseif strDBType = "sqlserver" then
		if strSQL_Server = "SQL7" then
			strUpdateSql = strSql(SQL7)
		else
			strUpdateSql = strSql(SQL6)
		end if
	elseif strDBType = "mysql" then
		strUpdateSql = strSql(MySql)
	end if

	my_Conn.Execute strUpdateSql

	SpecialErrors = 0
	Response.Write("<table cellspacing=""1"" cellpadding=""5"" width=""50%"" class=""tc"">" & strLE)
	for counter = 0 to my_Conn.Errors.Count -1
		ConnErrorNumber      = Err.Number
		ConnErrorDescription = my_Conn.Errors(counter).Description

		if ConnErrorNumber <> 0 then

			if ConnErrorNumber = -2147217900 then
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			elseif (instr(1,my_Conn.Errors(counter).Description,"Table",1) > 0) and (instr(1,my_Conn.Errors(counter).Description,"does not exist",1) > 0) and (instr(1,strUpdateSql,"DROP TABLE",1) > 0) then
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			elseif strDBType = "mysql" and instr(my_Conn.Errors(counter).Description, "already exists") > 0 then
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""orange"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intWarnings = intWarnings + 1
				SpecialErrors = 1
			else
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Error: " & Hex(ConnErrorNumber) & "</b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & ConnErrorDescription & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				Response.Write("<tr>" & strLE)
				Response.Write("<td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>strUpdateSql: </b></span></td>" & strLE)
				Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strUpdateSql & "</span></td>" & strLE)
				Response.Write("</tr>" & strLE)
				intCriticalErrors = intCriticalErrors + 1
				SpecialErrors = 1
			end if
		end if
	next
	if SpecialErrors = 0 then
		Response.Write("<tr>" & strLE)
		Response.Write("<td bgColor=""green"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></span></td>" & strLE)
		Response.Write("<td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & strOkMessage & "</span></td>" & strLE)
		Response.Write("</tr>" & strLE)
	end if

	Response.Write("</table>" & strLE)
	Response.Flush

	my_Conn.Errors.Clear
	Err.Clear
	on error goto 0
end sub

function CreateConnectionString(strConn, UserName, Password)
	'## strConnString = "driver={SQL Server};server=YYYY;uid=XX;pwd=XXXX;database=ZZZZZZ"
	'## strConnString = "driver={SQL Server};server=YYYY;USER ID=XX;PASSWORD=XXXX;database=ZZZZZZ"

	Dim TempConnString
	Dim uidTagStart(1), uidTagEnd
	Dim pwdTagStart(1), pwdTagEnd
	Dim uidTagStartPos, uidTagEndPos
	Dim pwdTagStartPos, pwdTagEndPos
	Dim blnUIDok, blnPWDok

	uidTagStart(0) = "UID="
	uidTagStart(1) = "USER ID="
	uidTagEnd      = ";"
	pwdTagStart(0) = "PWD="
	pwdTagStart(1) = "PASSWORD="
	pwdTagEnd      = ";"

	TempConnString = strConn
	blnUIDok = FALSE
	blnPWDok = FALSE

	for Counter = 0 to Ubound(uidTagStart)

		uidTagStartPos = InStr(1, UCase(TempConnString), UCase(uidTagStart(Counter)), 1)
		if uidTagStartPos > 0 then
			uidTagEndPos = InStr(uidTagStartPos, UCase(TempConnString), UCase(uidTagEnd), 1)
		else
			uidTagEndPos = 0
		end if

		if (uidTagStartpos > 0) and (uidTagEndPos > 0) then
			TempConnString = Left(TempConnString, (uidTagStartPos + len(uidTagStart(Counter))-1)) & UserName & Right(TempConnString, (len(TempConnString) - uidTagEndPos) + 1)
			blnUIDok = TRUE
		end if

		pwdTagStartPos = InStr(1, TempConnString, pwdTagStart(Counter), 1)
		if pwdTagStartPos > 0 then
			pwdTagEndPos = InStr(pwdTagStartPos, TempConnString, pwdTagEnd, 1)
		else
			pwdTagEndPos = 0
		end if

		if (pwdTagStartpos > 0) and (pwdTagEndPos > 0) then
			TempConnString = Left(TempConnString, (pwdTagStartPos + len(pwdTagStart(Counter))-1)) & Password & Right(TempConnString, (len(TempConnString) - pwdTagEndPos) + 1)
			blnPWDok = TRUE
		end if

	next
	if blnUIDok and blnPWDok then
		CreateConnectionString = TempConnString
	else
		CreateConnectionString = "<Error>"
	end if
end function

Sub TransferOldConfig

	on error resume next

	strSql = "SELECT * FROM " & strTablePrefix & "CONFIG WHERE CONFIG_ID = " & 1

	set rs = my_Conn.Execute(strSql)

	Response.Write("<table cellspacing=""0"" cellpadding=""5"" width=""50%"" class=""tc"">")

	UpdateErrorCode = UpdateErrorCheck()

	if UpdateErrorCode = 0 then
		if not(rs.eof) then
			strDummy = SetConfigValue(0,"STRVERSION", rs("C_STRVERSION"))
			strDummy = SetConfigValue(0,"STRFORUMTITLE", rs("C_STRFORUMTITLE"))
			strDummy = SetConfigValue(0,"STRCOPYRIGHT" , rs("C_STRCOPYRIGHT"))
			strDummy = SetConfigValue(0,"STRTITLEIMAGE", rs("C_STRTITLEIMAGE"))
			strDummy = SetConfigValue(0,"STRHOMEURL", rs("C_STRHOMEURL"))
			strDummy = SetConfigValue(0,"STRFORUMURL", rs("C_STRFORUMURL"))
			strDummy = SetConfigValue(0,"STRAUTHTYPE", rs("C_STRAUTHTYPE"))
			strDummy = SetConfigValue(0,"STRSETCOOKIETOFORUM", rs("C_STRSETCOOKIETOFORUM"))
			strDummy = SetConfigValue(0,"STREMAIL", rs("C_STREMAIL"))
			strDummy = SetConfigValue(0,"STRUNIQUEEMAIL", rs("C_STRUNIQUEEMAIL"))
			strDummy = SetConfigValue(0,"STRMAILMODE", rs("C_STRMAILMODE"))
			strDummy = SetConfigValue(0,"STRMAILSERVER", rs("C_STRMAILSERVER"))
			strDummy = SetConfigValue(0,"STRSENDER", rs("C_STRSENDER"))
			strDummy = SetConfigValue(0,"STRDATETYPE", rs("C_STRDATETYPE"))
			strDummy = SetConfigValue(0,"STRTIMETYPE", rs("C_STRTIMETYPE"))
			strDummy = SetConfigValue(0,"STRTIMEADJUSTLOCATION", rs("C_STRTIMEADJUSTLOCATION"))
			strDummy = SetConfigValue(0,"STRTIMEADJUST", rs("C_STRTIMEADJUST"))
			strDummy = SetConfigValue(0,"STRMOVETOPICMODE", rs("C_STRMOVETOPICMODE"))
			strDummy = SetConfigValue(0,"STRPRIVATEFORUMS", rs("C_STRPRIVATEFORUMS"))
			strDummy = SetConfigValue(0,"STRSHOWMODERATORS", rs("C_STRSHOWMODERATORS"))
			strDummy = SetConfigValue(0,"STRSHOWRANK", rs("C_STRSHOWRANK"))
			strDummy = SetConfigValue(0,"STRHIDEEMAIL", rs("C_STRHIDEEMAIL"))
			strDummy = SetConfigValue(0,"STRIPLOGGING", rs("C_STRIPLOGGING"))
			strDummy = SetConfigValue(0,"STRALLOWFORUMCODE", rs("C_STRALLOWFORUMCODE"))
			strDummy = SetConfigValue(0,"STRIMGINPOSTS", rs("C_STRIMGINPOSTS") )
			strDummy = SetConfigValue(0,"STRALLOWHTML", rs("C_STRALLOWHTML"))
			strDummy = SetConfigValue(0,"STRSECUREADMIN", rs("C_STRSECUREADMIN"))
			strDummy = SetConfigValue(0,"STRNOCOOKIES", rs("C_STRNOCOOKIES"))
			strDummy = SetConfigValue(0,"STREDITEDBYDATE", rs("C_STREDITEDBYDATE"))
			strDummy = SetConfigValue(0,"STRHOTTOPIC", rs("C_STRHOTTOPIC"))
			strDummy = SetConfigValue(0,"INTHOTTOPICNUM", rs("C_INTHOTTOPICNUM"))
			strDummy = SetConfigValue(0,"STRHOMEPAGE", rs("C_STRHOMEPAGE"))
			strDummy = SetConfigValue(0,"STRAIM", rs("C_STRAIM"))
			strDummy = SetConfigValue(0,"STRYAHOO", rs("C_STRYAHOO"))
			strDummy = SetConfigValue(0,"STRICQ", rs("C_STRICQ"))
			strDummy = SetConfigValue(0,"STRICONS", rs("C_STRICONS"))
			strDummy = SetConfigValue(0,"STRGFXBUTTONS", rs("C_STRGFXBUTTONS"))
			strDummy = SetConfigValue(0,"STRBADWORDFILTER", rs("C_STRBADWORDFILTER"))
			strDummy = SetConfigValue(0,"STRBADWORDS", rs("C_STRBADWORDS"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTFACE", rs("C_STRDEFAULTFONTFACE"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTSIZE ", rs("C_STRDEFAULTFONTSIZE"))
			strDummy = SetConfigValue(0,"STRHEADERFONTSIZE", rs("C_STRHEADERFONTSIZE"))
			strDummy = SetConfigValue(0,"STRFOOTERFONTSIZE", rs("C_STRFOOTERFONTSIZE"))
			strDummy = SetConfigValue(0,"STRPAGEBGCOLOR", rs("C_STRPAGEBGCOLOR"))
			strDummy = SetConfigValue(0,"STRDEFAULTFONTCOLOR", rs("C_STRDEFAULTFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRLINKCOLOR", rs("C_STRLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRLINKTEXTDECORATION", rs("C_STRLINKTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRVISITEDLINKCOLOR", rs("C_STRVISITEDLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRVISITEDTEXTDECORATION", rs("C_STRVISITEDTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRACTIVELINKCOLOR", rs("C_STRACTIVELINKCOLOR"))
			strDummy = SetConfigValue(0,"STRHOVERFONTCOLOR", rs("C_STRHOVERFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRHOVERTEXTDECORATION", rs("C_STRHOVERTEXTDECORATION"))
			strDummy = SetConfigValue(0,"STRHEADCELLCOLOR", rs("C_STRHEADCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRHEADFONTCOLOR", rs("C_STRHEADFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRCATEGORYCELLCOLOR", rs("C_STRCATEGORYCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRCATEGORYFONTCOLOR", rs("C_STRCATEGORYFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMFIRSTCELLCOLOR", rs("C_STRFORUMFIRSTCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMCELLCOLOR", rs("C_STRFORUMCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRALTFORUMCELLCOLOR", rs("C_STRALTFORUMCELLCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMFONTCOLOR", rs("C_STRFORUMFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRFORUMLINKCOLOR", rs("C_STRFORUMLINKCOLOR"))
			strDummy = SetConfigValue(0,"STRTABLEBORDERCOLOR", rs("C_STRTABLEBORDERCOLOR"))
			strDummy = SetConfigValue(0,"STRPOPUPTABLECOLOR", rs("C_STRPOPUPTABLECOLOR"))
			strDummy = SetConfigValue(0,"STRPOPUPBORDERCOLOR", rs("C_STRPOPUPBORDERCOLOR"))
			strDummy = SetConfigValue(0,"STRNEWFONTCOLOR", rs("C_STRNEWFONTCOLOR"))
			strDummy = SetConfigValue(0,"STRTOPICWIDTHLEFT", rs("C_STRTOPICWIDTHLEFT"))
			strDummy = SetConfigValue(0,"STRTOPICWIDTHRIGHT", rs("C_STRTOPICWIDTHRIGHT"))
			strDummy = SetConfigValue(0,"STRTOPICNOWRAPLEFT", rs("C_STRTOPICNOWRAPLEFT"))
			strDummy = SetConfigValue(0,"STRTOPICNOWRAPRIGHT", rs("C_STRTOPICNOWRAPRIGHT"))
			strDummy = SetConfigValue(0,"STRRANKADMIN", rs("C_STRRANKADMIN"))
			strDummy = SetConfigValue(0,"STRRANKMOD", rs("C_STRRANKMOD"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL0", rs("C_STRRANKLEVEL0"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL1", rs("C_STRRANKLEVEL1"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL2", rs("C_STRRANKLEVEL2"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL3", rs("C_STRRANKLEVEL3"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL4", rs("C_STRRANKLEVEL4"))
			strDummy = SetConfigValue(0,"STRRANKLEVEL5", rs("C_STRRANKLEVEL5"))
			strDummy = SetConfigValue(0,"STRRANKCOLORADMIN", rs("C_STRRANKCOLORADMIN"))
			strDummy = SetConfigValue(0,"STRRANKCOLORMOD", rs("C_STRRANKCOLORMOD"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR0", rs("C_STRRANKCOLOR0"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR1", rs("C_STRRANKCOLOR1"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR2", rs("C_STRRANKCOLOR2"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR3", rs("C_STRRANKCOLOR3"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR4", rs("C_STRRANKCOLOR4"))
			strDummy = SetConfigValue(0,"STRRANKCOLOR5", rs("C_STRRANKCOLOR5"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL0", rs("C_INTRANKLEVEL0"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL1", rs("C_INTRANKLEVEL1"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL2", rs("C_INTRANKLEVEL2"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL3", rs("C_INTRANKLEVEL3"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL4", rs("C_INTRANKLEVEL4"))
			strDummy = SetConfigValue(0,"INTRANKLEVEL5", rs("C_INTRANKLEVEL5"))
			strDummy = SetConfigValue(0,"STRSIGNATURES", rs("C_STRSIGNATURES") )
			strDummy = SetConfigValue(0,"STRSHOWSTATISTICS", rs("C_STRSHOWSTATISTICS"))
			strDummy = SetConfigValue(0,"STRSHOWIMAGEPOWEREDBY", rs("C_STRSHOWIMAGEPOWEREDBY"))
			strDummy = SetConfigValue(0,"STRLOGONFORMAIL", rs("C_STRLOGONFORMAIL"))
			strDummy = SetConfigValue(0,"STRSHOWPAGING", rs("C_STRSHOWPAGING"))
			strDummy = SetConfigValue(0,"STRSHOWTOPICNAV", rs("C_STRSHOWTOPICNAV"))
			strDummy = SetConfigValue(0,"STRPAGESIZE", rs("C_STRPAGESIZE"))
			strDummy = SetConfigValue(0,"STRPAGENUMBERSIZE", rs("C_STRPAGENUMBERSIZE"))
			strDummy = SetConfigValue(0,"STRFULLNAME", rs("C_STRFULLNAME"))
			strDummy = SetConfigValue(0,"STRPICTURE", rs("C_STRPICTURE"))
			strDummy = SetConfigValue(0,"STRSEX", rs("C_STRSEX"))
			strDummy = SetConfigValue(0,"STRCITY", rs("C_STRCITY"))
			strDummy = SetConfigValue(0,"STRSTATE", rs("C_STRSTATE"))
			strDummy = SetConfigValue(0,"STRAGE", rs("C_STRAGE"))
			strDummy = SetConfigValue(0,"STRCOUNTRY", rs("C_STRCOUNTRY"))
			strDummy = SetConfigValue(0,"STROCCUPATION", rs("C_STROCCUPATION"))
			strDummy = SetConfigValue(0,"STRHOMEPAGE", rs("C_STRHOMEPAGE"))
			strDummy = SetConfigValue(0,"STRFAVLINKS", rs("C_STRFAVLINKS"))
			strDummy = SetConfigValue(0,"STRBIO", rs("C_STRBIO"))
			strDummy = SetConfigValue(0,"STRHOBBIES", rs("C_STRHOBBIES"))
			strDummy = SetConfigValue(0,"STRLNEWS", rs("C_STRLNEWS"))
			strDummy = SetConfigValue(0,"STRQUOTE", rs("C_STRQUOTE"))
			strDummy = SetConfigValue(0,"STRMARSTATUS", rs("C_STRMARSTATUS"))
			strDummy = SetConfigValue(0,"STRRECENTTOPICS", rs("C_STRRECENTTOPICS"))
			strDummy = SetConfigValue(0,"STRNTGROUPS", rs("C_STRNTGROUPS"))
			strDummy = SetConfigValue(0,"STRAUTOLOGON", rs("C_STRAUTOLOGON"))
			strDummy = SetConfigValue(0,"STRMOVENOTIFY", "1")
			strDummy = SetConfigValue(0,"STRSUBSCRIPTION", "1")
			strDummy = SetConfigValue(0,"STRMODERATION", "1")

			Response.Write("<tr><td bgColor=""green"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></span></td><td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2""> Config values transferred to new table</span></td></tr>")
		else
			Response.Write("<tr><td bgColor=orange align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Upgraded: </b></span></td><td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2""> No existing config values found</span></td></tr>")
		end if
	else
		Response.Write("<tr><td bgColor=""red"" align=""left"" width=""30%""><span face=""Verdana, Arial, Helvetica"" size=""2""><b>Critical error: code: </b></span></td><td bgColor=""#9FAFDF"" align=""left""><span face=""Verdana, Arial, Helvetica"" size=""2"">" & Hex(UpdateErrorCode) & " while trying to tranfer the existing config values to the new table</span></td></tr>")
		intCriticalErrors = intCriticalErrors + 1
	end if
	Response.Write("</table>")

	rs.close
	set rs = nothing

	on error goto 0
end sub

Sub UpDateAccessFields(pOldversion)

	if pOldversion <= 5  then

		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_MODERATION = 0 WHERE (CAT_MODERATION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_SUBSCRIPTION = 0 WHERE (CAT_SUBSCRIPTION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "CATEGORY SET CAT_ORDER = 1 WHERE (CAT_ORDER Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_L_ARCHIVE = '' WHERE (F_L_ARCHIVE Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_ARCHIVE_SCHED = 30 WHERE (F_ARCHIVE_SCHED Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_L_DELETE = '' WHERE (F_L_DELETE Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_DELETE_SCHED = 365 WHERE (F_DELETE_SCHED Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_MODERATION = 0 WHERE (F_MODERATION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_SUBSCRIPTION = 0 WHERE (F_SUBSCRIPTION Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "FORUM SET F_ORDER = 1 WHERE (F_ORDER Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "TOPICS SET T_ARCHIVE_FLAG = 1 WHERE (T_ARCHIVE_FLAG Is Null)")
		my_Conn.execute ("UPDATE " & strTablePrefix & "REPLY SET R_STATUS = 0 WHERE (R_STATUS Is Null)")

		on error resume next
		my_Conn.execute("ALTER TABLE " & strTablePrefix & "TOPICS DROP COLUMN C_STRMOVENOTIFY")
		on error goto 0
	end if
end sub

function SetConfigValue(bUpdate, fVariable, fValue)

	' bUpdate = 1 : if it exists then overwrite with new values
	' bUpdate = 0 : if it exists then leave unchanged

	Dim strSql

	strSql = "SELECT C_VARIABLE FROM " & strTablePrefix & "CONFIG_NEW " &_
		 " WHERE C_VARIABLE = '" & fVariable & "' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if (rs.EOF or rs.BOF) then '## New config-value
		SetConfigValue = "added"
		my_conn.execute ("INSERT INTO " & strTablePrefix & "CONFIG_NEW (C_VALUE,C_VARIABLE) VALUES ('" & fValue & "' , '" & fVariable & "')")
	else
		if bUpdate <> 0 then
			SetConfigValue = "updated"
			my_conn.execute ("UPDATE " & strTablePrefix & "CONFIG_NEW SET C_VALUE = '" & fValue & "' WHERE C_VARIABLE = '" & fVariable &"'")
		else ' not changed
			SetConfigValue = "unchanged"
		end if
	end if

	rs.close
	set rs = nothing
end function

function SetBadWordValue(bUpdate, fVariable, fValue)

	' bUpdate = 1 : if it exists then overwrite with new values
	' bUpdate = 0 : if it exists then leave unchanged

	Dim strSql

	strSql = "SELECT B_BADWORD FROM " & strFilterTablePrefix & "BADWORDS " &_
		 " WHERE B_BADWORD = '" & fVariable & "' "

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.open strSql, my_Conn

	if (rs.EOF or rs.BOF) then '## New Badword
		SetBadWordValue = "added"
		my_conn.execute ("INSERT INTO " & strFilterTablePrefix & "BADWORDS (B_REPLACE,B_BADWORD) VALUES ('" & chkString(fValue,"sqlstring") & "' , '" & fVariable & "')")
	else
		if bUpdate <> 0 then
			SetBadWordValue = "updated"
			my_conn.execute ("UPDATE " & strFilterTablePrefix & "BADWORDS SET B_REPLACE = '" & chkString(fValue,"sqlstring") & "' WHERE B_BADWORD = '" & fVariable &"'")
		else ' not changed
			SetBadWordValue = "unchanged"
		end if
	end if

	rs.close
	set rs = nothing
end function

function doublenum(fNum)
	if fNum > 9 then
		doublenum = fNum
	else
		doublenum = "0" & fNum
	end if
end function

function DateToStr(dtDateTime)
	if not isDate(dtDateTime) then
		dtDateTime = strToDate(dtDateTime)
	end if

	DateToStr = year(dtDateTime) & doublenum(Month(dtdateTime)) & doublenum(Day(dtdateTime)) & doublenum(Hour(dtdateTime)) & doublenum(Minute(dtdateTime)) & doublenum(Second(dtdateTime)) & ""
end function

Function IsValidString(sValidate)
	Dim sInvalidChars
	Dim bTemp
	Dim i
	' Disallowed characters
	sInvalidChars = "!#$%^&*()=+{}[]|\;:/?>,<'"
	for i = 1 To Len(sInvalidChars)
		if InStr(sValidate, Mid(sInvalidChars, i, 1)) > 0 then bTemp = True
		if bTemp then Exit For
	next
	for i = 1 to Len(sValidate)
		if Asc(Mid(sValidate, i, 1)) = 160 then bTemp = True
		if bTemp then Exit For
	next

	' extra checks
	' no two consecutive dots or spaces
	if not bTemp then
		bTemp = InStr(sValidate, "..") > 0
	end if
	if not bTemp then
		bTemp = InStr(sValidate, "  ") > 0
	end if
	if not bTemp then
		bTemp = (len(sValidate) <> len(Trim(sValidate)))
	end if 'Addition for leading and trailing spaces

	' if any of the above are true, invalid string
	IsValidString = Not bTemp
End Function

function CheckSelected(ByVal chkval1, chkval2)
	if IsNumeric(chkval1) then chkval1 = cLng(chkval1)
	if (chkval1 = chkval2) then
		CheckSelected = " selected"
	else
		CheckSelected = ""
	end if
end function

function HTMLEncode(pString)
	fString    = trim(pString)
	if fString = "" or IsNull(fString) then fString = " "
	fString    = replace(fString, ">", "&gt;")
	fString    = replace(fString, "<", "&lt;")
	HTMLEncode = fString
end function

function chkString(pString,fField_Type) '## Types - SQLString
	fString = trim(pString)
	if fString = "" or isNull(fString) then
		fString = " "
	end if
	select case fField_Type
		case "SQLString"
			fString = Replace(fString, "'", "''")
			if strDBType = "mysql" then
				fString = Replace(fString, "\0", "\\0")
				fString = Replace(fString, "\'", "\\'")
				fString = Replace(fString, "\""", "\\""")
				fString = Replace(fString, "\b", "\\b")
				fString = Replace(fString, "\n", "\\n")
				fString = Replace(fString, "\r", "\\r")
				fString = Replace(fString, "\t", "\\t")
				fString = Replace(fString, "\z", "\\z")
				fString = Replace(fString, "\%", "\\%")
				fString = Replace(fString, "\_", "\\_")
			end if
			fString = HTMLEncode(fString)
			chkString = fString
			exit function
	end select
	chkString = fString
end function
%>
