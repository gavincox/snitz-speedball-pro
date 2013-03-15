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
<!--#INCLUDE FILE="inc_sha256.asp"-->
<!--#INCLUDE FILE="inc_header_short.asp" -->
<%
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login_short.asp?target=" & scriptname(ubound(scriptname))
end if
strRqMethod = trim(chkString(Request.QueryString("method"),"SQLString"))
intUsernameID = trim(chkString(Request.QueryString("N_ID"),"SQLString"))
if intUsernameID <> "" then
	if isNumeric(intUsernameID) <> True then intUsernameID = "0"
end if
strPageSize = 10
mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)
Response.Write "<script type=""text/javascript"">" & strLE & _
	"<!-- " & vbNewLine & _
	"function jumpToPage(s) {location.href = s.options[s.selectedIndex].value}" & strLE & _
	"// -->" & strLE & _
	"</script>" & strLE
Select Case strRqMethod
	Case "Add"
		if Request.Form("Method_Type") = "Write_Configuration" then
			Err_Msg = ""
			if not IsValidString(trim(Request.Form("strUserName"))) then Err_Msg = Err_Msg & "<li>None of the following characters can be used in the username  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
			txtUserName = chkString(Request.Form("strUserName"),"SQLString")
			if txtUserName = " " then Err_Msg = Err_Msg & "<li>You Must Enter a UserName to filter.</li>"
			if (Instr(txtUserName, "  ") > 0 ) then Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the UserName.</li>"
			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "INSERT INTO " & strFilterTablePrefix & "NAMEFILTER ("
				strSql = strSql & "N_NAME"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & txtUserName & "'"
				strSql = strSql & ")"
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				Application.Lock
				Application(strCookieURL & "STRFILTERUSERNAMES") = ""
				Application.UnLock
				Response.Write "<p class=""c""><span class=""dff hfs"">UserName Added!</span></p>" & strLE & _
					"<meta http-equiv=""Refresh"" content=""1; URL=admin_config_namefilter.asp"">" & strLE & _
					"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""admin_config_namefilter.asp"">Back To UserName Filter Configuration</span></a></p>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
					"<table class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
			end if
		end if
	Case "Delete"
		if Request.Form("Method_Type") = "Delete_UserName" then
			'## Forum_SQL - Delete UserName from NameFilter table
			strSql = "DELETE FROM " & strFilterTablePrefix & "NAMEFILTER "
			strSql = strSql & " WHERE N_ID = " & Request.Form("N_ID")
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
			Application.Lock
			Application(strCookieURL & "STRFILTERUSERNAMES") = ""
			Application.UnLock
			Response.Write "<p class=""c""><span class=""dff hfs""><b>UserName Deleted!</b></span></p>" & strLE & _
				"<meta http-equiv=""Refresh"" content=""1; URL=admin_config_namefilter.asp"">" & strLE & _
				"<p class=""c""><span class=""dff dfs""><a href=""admin_config_namefilter.asp"">Back To UserName Filter Configuration</span></a></p>" & strLE
		else
			Response.Write "<form action=""admin_config_namefilter.asp?method=Delete"" method=""post"" id=""DeleteUserName"" name=""DeleteUserName"">" & strLE & _
				"<input type=""hidden"" name=""Method_Type"" value=""Delete_UserName"">" & strLE & _
				"<input type=""hidden"" name=""N_ID"" value=""" & intUsernameID & """>" & strLE & _
				"<p class=""c""><span class=""dff hfs""><b>Are you sure?</b></span></p>" & strLE & _
				"<p class=""c""><input type=""submit"" class=""button"" value=""Yes"" id=""submit1"" name=""submit1"">&nbsp;<input type=""button"" class=""button"" value="" No "" onClick=""history.go(-1);""></p>" & strLE & _
				"</form>" & strLE
		end if
	Case "Edit"
		if Request.Form("Method_Type") = "Write_Configuration" then
			Err_Msg = ""
			if not IsValidString(trim(Request.Form("strUserName"))) then
				Err_Msg = Err_Msg & "<li>None of the following characters can be used in the username  !#$%^&*()=+{}[]|\;:/?>,<' </li>"
			end if
			txtUserName = chkString(Request.Form("strUserName"),"SQLString")
			if txtUserName = " " then
				Err_Msg = Err_Msg & "<li>You Must Enter a UserName.</li>"
			end if
			if (Instr(txtUserName, "  ") > 0 ) then
				Err_Msg = Err_Msg & "<li>Two or more consecutive spaces are not allowed in the UserName.</li>"
			end if
			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strFilterTablePrefix & "NAMEFILTER "
				strSql = strSql & " SET N_NAME = '" & txtUserName & "'"
				strSql = strSql & " WHERE N_ID = " & Request.Form("N_ID")
				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
				Application.Lock
				Application(strCookieURL & "STRFILTERUSERNAMES") = ""
				Application.UnLock
				Response.Write "<p class=""c""><span class=""dff hfs"">UserName Filter Updated!</span></p>" & strLE & _
					"<meta http-equiv=""Refresh"" content=""1; URL=admin_config_namefilter.asp"">" & strLE & _
					"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""admin_config_namefilter.asp"">Back To UserName Filter Configuration</span></a></p>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
					"<table class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></p>" & strLE
			end if
		else
			'## Forum_SQL - Get UserName from DB
			strSql = "SELECT N_ID, N_NAME "
			strSql = strSql & " FROM " & strFilterTablePrefix & "NAMEFILTER "
			strSql = strSql & " WHERE N_ID = " & intUsernameID
			set rs = my_Conn.Execute (strSql)
		        TxtUserName = rs("N_NAME")
		        intN_ID = rs("N_ID")
			rs.close
			set rs = nothing
			Response.Write "<form action=""admin_config_namefilter.asp?method=Edit"" method=""post"" id=""UpdateUserName"" name=""UpdateUserName"">" & strLE & _
				"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
				"<input type=""hidden"" name=""N_ID"" value=""" & intN_ID & """>" & strLE & _
				"<table class=""admin"">" & strLE & _
				"<th><b>Edit UserName</b></td>" & strLE & _
				"</tr>" & strLE & _
				"<th><b>Username</b></th>" & strLE & _
				"</tr>" & strLE & _
				"<tr class=""c"">" & strLE & _
				"<td><input size=""25"" maxLength=""25"" name=""strUserName"" value=""" & TxtUserName & """ tabindex=""1""></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr class=""c"">" & strLE & _
				"<td><input type=""submit"" class=""button"" value=""Update"" id=""submit1"" name=""submit1"" tabindex=""2""> <input type=""reset"" class=""button"" value=""Reset"" id=""reset1"" name=""reset1""></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</form>" & strLE & _
				"<p class=""c""><a href=""admin_config_namefilter.asp"">Back To UserName Filter Configuration</a></p>" & strLE
		end if
	Case Else
		'## Forum_SQL - Get UserNames from DB
		strSql = "SELECT N_ID, N_NAME "
		strSql2 = " FROM " & strFilterTablePrefix & "NAMEFILTER "
		strSql3 = " ORDER BY N_NAME ASC "
		if strDBType = "mysql" then 'MySql specific code
			if mypage > 1 then
				OffSet = cLng((mypage - 1) * strPageSize)
				strSql4 = " LIMIT " & OffSet & ", " & strPageSize & " "
			end if
			'## Forum_SQL - Get the total pagecount
			strSql1 = "SELECT COUNT(N_ID) AS PAGECOUNT "
			set rsCount = my_Conn.Execute(strSql1 & strSql2)
			iPageTotal = rsCount(0).value
			rsCount.close
			set rsCount = nothing
			If iPageTotal > 0 then
				maxpages = (iPageTotal \ strPageSize )
				if iPageTotal mod strPageSize <> 0 then
					maxpages = maxpages + 1
				end if
				if iPageTotal < (strPageSize + 1) then
					intGetRows = iPageTotal
				elseif (mypage * strPageSize) > iPageTotal then
					intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
				else
					intGetRows = strPageSize
				end if
			else
				iPageTotal = 0
				maxpages = 0
			end if
			if iPageTotal > 0 then
				set rs = Server.CreateObject("ADODB.Recordset")
				rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
					arrUsernameData = rs.GetRows(intGetRows)
					iUsernameCount = UBound(arrUsernameData, 2)
				rs.close
				set rs = nothing
			else
				iUsernameCount = ""
			end if
		else 'end MySql specific code
			set rs = Server.CreateObject("ADODB.Recordset")
			rs.cachesize = strPageSize
			rs.open strSql & strSql2 & strSql3, my_Conn, adOpenStatic
				If not (rs.EOF or rs.BOF) then
					rs.movefirst
					rs.pagesize = strPageSize
					rs.absolutepage = mypage '**
					maxpages = cLng(rs.pagecount)
					arrUsernameData = rs.GetRows(strPageSize)
					iUsernameCount = UBound(arrUsernameData, 2)
				else
					iUsernameCount = ""
				end if
			rs.Close
			set rs = nothing
		end if
		Response.Write "<p class=""c""><span class=""dff hfs""><b>UserName Filter Configuration</b></span></p>" & strLE
		Response.Write "<form action=""admin_config_namefilter.asp?method=Add"" method=""post"" id=""Add"" name=""Add"">" & strLE & _
			"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
			"<table class=""admin"">" & strLE & _
			"<th><b>UserName</b></th>" & strLE
		if maxpages > 1 then
			Call DropDownPaging()
		else
			Response.Write "<th><b>&nbsp;</b></th>" & strLE
		end if
		Response.Write "</tr>" & strLE
		if iUsernameCount = "" then  '## No Badwords found in DB
			Response.Write "<tr>" & strLE & _
				"<td colspan=""2""><b>No UserNames Found</b></td>" & strLE & _
				"</tr>" & strLE
		else
			nN_ID = 0
			nN_NAME = 1
			rec = 1
			'intI = 0
			for iUsername = 0 to iUsernameCount
				if (rec = strPageSize + 1) then exit for
				Username_ID = arrUsernameData(nN_ID, iUsername)
				Username_Name = arrUsernameData(nN_NAME, iUsername)
				'if intI = 1 then
				'	CColor = strAltForumCellColor
				'else
				'	CColor = strForumCellColor
				'end if
				Response.Write "<tr>" & strLE & _
					"<td>" & Username_Name & "</td>" & strLE & _
					"<td><a href=""admin_config_namefilter.asp?method=Edit&N_ID=" & Username_ID & """>" & _
					getCurrentIcon(strIconPencil,"Edit UserName","class=""vam""") & "</a>&nbsp;<a href=""admin_config_namefilter.asp?method=Delete&N_ID=" & Username_ID & """>" & _
					getCurrentIcon(strIconTrashcan,"Delete UserName","class=""vam""") & "</a></td>" & strLE & _
					"</tr>" & strLE
				rec = rec + 1
				'intI = intI + 1
				'if intI = 2 then intI = 0
			next
		end if
		Response.Write "<tr class=""c"">" & strLE & _
			"<td><input size=""25"" maxLength=""25"" name=""strUserName"" value=""" & TxtUserName & """ tabindex=""1""></td>" & strLE & _
			"<td><input class=""button"" value=""Add"" type=""submit"" tabindex=""2""></a></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE & _
			"</form>" & strLE
End Select
Call WriteFooterShort
Response.End
sub DropDownPaging()
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		Response.Write "<td class=""ccc vam""><span class=""dff ffs cfc"">" & strLE & _
			"<b>Page</b>&nbsp;<select style=""font-size:9px"" name=""whichpage"" size=""1"" onchange=""jumpToPage(this)"">" & strLE
		for counter = 1 to maxpages
			ref = "admin_config_namefilter.asp?whichpage=" & counter
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & ref & """>" & counter & "</option>" & strLE
			else
				Response.Write "<option value=""" & ref & """ selected>" & counter & "</option>" & strLE
			end if
		next
		Response.Write "</select>&nbsp;<b>of " & maxpages & "</b></span></td>" & strLE
	end if
end sub
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
	if not bTemp then bTemp = InStr(sValidate, "..") > 0
	if not bTemp then bTemp = InStr(sValidate, "  ") > 0
	if not bTemp then bTemp = (len(sValidate) <> len(Trim(sValidate))) 'Addition for leading and trailing spaces
	' if any of the above are true, invalid string
	IsValidString = Not bTemp
End Function
%>
