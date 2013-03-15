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
	if Request.QueryString <> "" then
		Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname)) & "?" & Request.QueryString
	else
		Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
	end if
end if
Response.Write "<div id=""pre-content"">" & strLE & _
	"<div class=""breadcrumbs"">" & strLE & _
	getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""default.asp"">" & chkString(strForumTitle,"pagetitle") & "</a><br>" & strLE & _
	getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpen,"","class=""vam""") & "&nbsp;<a href=""admin_home.asp"">Admin&nbsp;Section</a><br>" & strLE & _
	getCurrentIcon(strIconBlank,"","class=""vam""") & getCurrentIcon(strIconBar,"","class=""vam""") & getCurrentIcon(strIconFolderOpenTopic,"","class=""vam""") & "&nbsp;<a href=""admin_config_groupcats.asp"">Group&nbsp;Categories&nbsp;Configuration</a>" & strLE & _
	"</div>" & strLE & _
	"<!-- /breadcrumbs -->" & strLE & _
	"<div class=""maxpages"">" & strLE & _
	"</div>" & strLE & _
	"<!-- /maxpages -->" & strLE & _
	"</div>" & strLE & _
	"<!-- /pre-content -->" & strLE & _
	"<br>" & strLE & strLE

Response.Write "<script type=""text/javascript"" src=""selectbox.js""></script>" & strLE

strRqMethod = Request.QueryString("method")

Select Case strRqMethod
	Case "Add"
		if Request.Form("Method_Type") = "Write_Configuration" then
			Err_Msg = ""

			txtGroupName = chkString(Request.Form("strGroupName"),"SQLString")
			txtGroupDescription = chkString(Request.Form("strGroupDescription"),"message")
			txtGroupIcon = chkString(Request.Form("strGroupIcon"),"SQLString")
			txtGroupTitleImage = chkString(Request.Form("strGroupTitleImage"),"SQLString")

			if trim(txtGroupName) = "" then
				Err_Msg = Err_Msg & "<li>You Must Enter a Name for your New Group.</li>"
			end if

			if trim(txtGroupDescription) = "" then
				Err_Msg = Err_Msg & "<li>You Must Enter a Description for your New Group.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "INSERT INTO " & strTablePrefix & "GROUP_NAMES ("
				strSql = strSql & "GROUP_NAME"
				strSql = strSql & ", GROUP_DESCRIPTION"
				strSql = strSql & ", GROUP_ICON"
				strSql = strSql & ", GROUP_IMAGE"
				strSql = strSql & ") VALUES ("
				strSql = strSql & "'" & txtGroupName & "'"
				strSql = strSql & ", '" & txtGroupDescription & "'"
				strSql = strSql & ", '" & txtGroupIcon & "'"
				strSql = strSql & ", '" & txtGroupTitleImage & "'"
				strSql = strSql & ")"

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				set rsCount = my_Conn.execute("SELECT MAX(GROUP_ID) AS maxGroupID FROM " & strTablePrefix & "GROUP_NAMES ")
				newGroupCategories rsCount("maxGroupId")
				set rsCount = nothing

				Response.Write "<p class=""c""><span class=""dff hfs"">New Group Added!</span></p>" & strLE & _
					"<meta http-equiv=""Refresh"" content=""2; URL=admin_config_groupcats.asp"">" & strLE & _
					"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></span></p>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
					"<table class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Enter Data</a></span></p>" & strLE
			end if
		else
			Response.Write "<form action=""admin_config_groupcats.asp?method=Add"" method=""post"" id=""Add"" name=""Add"">" & strLE & _
				"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
				"<table class=""admin"">" & strLE & _
				"<tr>" & strLE & _
				"<th colspan=""2""><b>Create A New Category Group</b></th>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw r""><b>New Group Name</b>&nbsp;</td>" & strLE & _
				"<td class=""l""><input maxLength=""50"" name=""strGroupName"" value="""" tabindex=""1"" size=""46""></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr class=""vat"">" & strLE & _
				"<td class=""nw r""><b>New Group Description</b>&nbsp;</td>" & strLE & _
				"<td class=""l""><textarea maxLength=""255"" rows=""5"" cols=""35"" name=""strGroupDescription"" tabindex=""2""></textarea></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw r""><b>New Group Icon</b>&nbsp;</td>" & strLE & _
				"<td class=""l""><input maxLength=""255"" name=""strGroupIcon"" value="""" tabindex=""3"" size=""46""></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw r""><b>New Group Title Image</b>&nbsp;</td>" & strLE & _
				"<td class=""l""><input maxLength=""255"" name=""strGroupTitleImage"" value="""" tabindex=""4"" size=""46""></td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""vat nw r""><b>Categories</b>&nbsp;</td>" & strLE
			strSql = "SELECT CAT_ID, CAT_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
			strSql = strSql & " ORDER BY CAT_NAME ASC "

			set rsCats = Server.CreateObject("ADODB.Recordset")
			rsCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			if rsCats.EOF then
				recCatCnt = ""
			else
				allCatData = rsCats.GetRows(adGetRowsRest)
				recCatCnt = UBound(allCatData,2)
				cCAT_ID = 0
				cCAT_NAME = 1
			end if

			rsCats.close
			set rsCats = nothing

			SelectSize = 6
			Response.Write "<td>" & strLE & _
					"<table class=""tnb"">" & strLE & _
					"<tr>" & strLE & _
					"<td class=""c""><b>Available</b><br>" & strLE & _
					"<select name=""GroupCatCombo"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & strLE
			'## Pick from list
			if recCatCnt <> "" then
				for iCat = 0 to recCatCnt
					CategoryCatID = allCatData(cCAT_ID,iCat)
					CategoryCatName = allCatData(cCAT_NAME,iCat)
					Response.Write 	"<option value=""" & CategoryCatID & """>" & ChkString(CategoryCatName,"display") & "</option>" & strLE
				next
			end if
			Response.Write "</select>" & strLE & _
				"</td>" & strLE & _
				"<td class=""vam c"" width=""15""><br>" & strLE & _
				"<a href=""javascript:moveAllOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemAll,"","class=""vam""") & "</a>" & strLE & _
				"<a href=""javascript:moveSelectedOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemove,"","class=""vam""") & "</a>" & strLE & _
				"<a href=""javascript:moveSelectedOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAdd,"","class=""vam""") & "</a>" & strLE & _
				"<a href=""javascript:moveAllOptions(document.Add.GroupCatCombo, document.Add.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAddAll,"","class=""vam""") & "</a>" & strLE & _
				"</td>" & strLE & _
				"<td class=""c""><b>Selected</b><br>" & strLE & _
				"<select name=""GroupCat"" size=""" & SelectSize & """ multiple tabindex=""15"" onDblClick=""moveSelectedOptions(document.Add.GroupCat, document.Add.GroupCatCombo, true, '')"">" & strLE & _
				"</select>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw c"" colspan=""2""><input class=""button"" value=""  Add  "" type=""submit"" tabindex=""5"" onclick=""selectAllOptions(document.Add.GroupCat);"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset"" tabindex=""6""></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</form>" & strLE & _
				"<p class=""c""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></p>" & strLE
		end if
	Case "Delete"
		if Request.Form("Method_Type") = "Delete_Category" then
			'## Forum_SQL
			strSql = "DELETE FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			strSql = "DELETE FROM " & strTablePrefix & "GROUPS "
			strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

               		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

			Response.Write "<p class=""c""><span class=""dff hfs""><b>Category Group Deleted!</b></span></p>" & strLE & _
				"<meta http-equiv=""Refresh"" content=""1; URL=admin_config_groupcats.asp"">" & strLE & _
				"<p class=""c""><span class=""dff dfs""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></span></p>" & strLE
		else
			'## Forum_SQL
			strSql = "SELECT GROUP_ID, GROUP_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID <> 1 "
			strSql = strSql & " AND GROUP_ID <> 2 "
			strSql = strSql & " ORDER BY GROUP_NAME ASC "

			Set rsgroups = Server.CreateObject("ADODB.Recordset")
			rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			If rsgroups.EOF then
				recGroupCount = ""
			Else
				allGroupData = rsgroups.GetRows(adGetRowsRest)
				recGroupCount = UBound(allGroupData, 2)
			End if

			rsgroups.Close
			Set rsgroups = Nothing

			Response.Write "<script type=""text/javascript"">" & strLE & _
				"<!-- " & vbNewLine & _
				"function confirmDelete(){" & strLE & _
				"var where_to= confirm(""Do you really want to Delete this Group Category?"", ""Yes"", ""No"");" & strLE & _
				"if (where_to)" & strLE & _
				"return true;" & strLE & _
				"else" & strLE & _
				"return false;" & strLE & _
				"}" & strLE & _
				"//-->" & strLE & _
				"</script>" & strLE

			Response.Write "<form action=""admin_config_groupcats.asp?method=Delete"" method=""post"" id=""Add"" name=""Add"">" & strLE & _
				"<input type=""hidden"" name=""Method_Type"" value=""Delete_Category"">" & strLE & _
				"<table class=""admin"">" & strLE & _
				"<tr>" & strLE & _
				"<th colspan=""2""><b>Delete Group Categories</b></th>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE
			if recGroupCount <> "" then
				Response.Write "<td class=""nw r""><b>Choose Group To Delete</b>&nbsp;</td>" & strLE & _
					"<td class=""l""><select name=""GroupID"" size=""1"">" & strLE
				for iGroup = 0 to recGroupCount
					Response.Write "<option value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0,iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & strLE
				next
				Response.Write "</select>" & strLE & _
					"</td>" & strLE & _
					"</tr>" & strLE & _
					"<tr>" & strLE & _
					"<td class=""nw c"" colspan=""2""><input class=""button"" value="" Delete "" type=""submit"" onClick=""return confirmDelete()"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset""></td>" & strLE
			else
				Response.Write "<td class=""nw c"" colspan=""2"">&nbsp;<b><i>No Groups Available To Delete</i></b>&nbsp;</td>" & strLE
			end if
			Response.Write "</tr>" & strLE & _
				"</table>" & strLE & _
				"</form>" & strLE & _
				"<p class=""c""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></p>" & strLE
		end if
	Case "Edit"
		if Request.Form("Method_Type") = "Write_Configuration" then
			txtGroupName = chkString(Request.Form("strGroupName"),"SQLString")
			txtGroupDescription = chkString(Request.Form("strGroupDescription"),"message")
			txtGroupIcon = chkString(Request.Form("strGroupIcon"),"SQLString")
			txtGroupTitleImage = chkString(Request.Form("strGroupTitleImage"),"SQLString")

			if trim(txtGroupName) = "" then
				Err_Msg = Err_Msg & "<li>You Must Enter a Name for your New Group.</li>"
			end if

			if trim(txtGroupDescription) = "" then
				Err_Msg = Err_Msg & "<li>You Must Enter a Description for your New Group.</li>"
			end if

			if Err_Msg = "" then
				'## Forum_SQL - Do DB Update
				strSql = "UPDATE " & strTablePrefix & "GROUP_NAMES "
				strSql = strSql & " SET GROUP_NAME = '" & txtGroupName & "'"
				strSql = strSql & ",    GROUP_DESCRIPTION = '" & txtGroupDescription & "'"
				strSql = strSql & ",    GROUP_ICON = '" & txtGroupIcon & "'"
				strSql = strSql & ",    GROUP_IMAGE = '" & txtGroupTitleImage & "'"
				strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GROUP_ID"))

				my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords

				updateGroupCategories(Request.Form("GROUP_ID"))

				Response.Write "<p class=""c""><span class=""dff hfs"">Category Group Updated!</span></p>" & strLE & _
					"<meta http-equiv=""Refresh"" content=""2; URL=admin_config_groupcats.asp"">" & strLE & _
					"<p class=""c""><span class=""dff hfs"">Congratulations!</span></p>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></span></p>" & strLE
			else
				Response.Write "<p class=""c""><span class=""dff hfs hlfc"">There Was A Problem With Your Details</span></p>" & strLE & _
					"<table class=""tc"">" & strLE & _
					"<tr>" & strLE & _
					"<td><span class=""dff dfs hlfc""><ul>" & Err_Msg & "</ul></span></td>" & strLE & _
					"</tr>" & strLE & _
					"</table>" & strLE & _
					"<p class=""c""><span class=""dff dfs""><a href=""JavaScript:history.go(-1)"">Go Back To Correct The Problem</a></span></p>" & strLE
			end if
		elseif Request.Form("Method_Type") = "Edit_Category" then
			if Request.Form("GroupID") <> "" then
				'## Forum_SQL
				strSql = "SELECT GROUP_ID, GROUP_NAME, GROUP_DESCRIPTION, GROUP_ICON, GROUP_IMAGE  "
				strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
				strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

				set rsGroups = Server.CreateObject("ADODB.Recordset")
				rsGroups.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

				if rsGroups.EOF then
					recGroupCnt = ""
				else
					allGroupData = rsGroups.GetRows(adGetRowsRest)
					recGroupCnt = UBound(allGroupData,2)
					gGROUP_ID = 0
					gGROUP_NAME = 1
					gGROUP_DESCRIPTION = 2
					gGROUP_ICON = 3
					gGROUP_IMAGE = 4
				end if

				rsGroups.close
				set rsGroups = nothing

				if recGroupCnt <> "" then
					txtGroupID = allGroupData(gGROUP_ID,0)
					txtGroupName = allGroupData(gGROUP_NAME,0)
					txtGroupDescription = allGroupData(gGROUP_DESCRIPTION,0)
					txtGroupIcon = allGroupData(gGROUP_ICON,0)
					txtGroupTitleImage = allGroupData(gGROUP_IMAGE,0)

					Response.Write "<form action=""admin_config_groupcats.asp?method=Edit"" method=""post"" id=""Edit"" name=""Edit"">" & strLE & _
						"<input type=""hidden"" name=""Method_Type"" value=""Write_Configuration"">" & strLE & _
						"<input type=""hidden"" name=""GROUP_ID"" value=""" & txtGroupID & """>" & strLE & _
						"<table class=""admin"">" & strLE & _
						"<tr>" & strLE & _
						"<th colspan=""2""><b>Edit Existing Category Group</b></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""nw r""><b>Group Name</b>&nbsp;</td>" & strLE & _
						"<td class=""l""><input maxLength=""50"" name=""strGroupName"" value=""" & txtGroupName & """ tabindex=""1"" size=""46""></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr class=""vat"">" & strLE & _
						"<td class=""nw r""><b>Group Description</b>&nbsp;</td>" & strLE & _
						"<td class=""l""><textarea rows=""5"" cols=""35"" name=""strGroupDescription"" maxLength=""255"" tabindex=""2"">" & txtGroupDescription & "</textarea></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""nw r""><b>Group Icon</b>&nbsp;</td>" & strLE & _
						"<td class=""l""><input maxLength=""255"" name=""strGroupIcon"" value=""" & txtGroupIcon & """ tabindex=""3"" size=""46""></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr class=""vam"">" & strLE & _
						"<td class=""nw r""><b>Group Title Image</b>&nbsp;</td>" & strLE & _
						"<td class=""l""><input maxLength=""255"" name=""strGroupTitleImage"" value=""" & txtGroupTitleImage & """ tabindex=""4"" size=""46""></td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""vat r nw ""><b>Categories</b>&nbsp;</td>" & strLE
					strSql = "SELECT CAT_ID, CAT_NAME "
					strSql = strSql & " FROM " & strTablePrefix & "CATEGORY "
					strSql = strSql & " ORDER BY CAT_NAME ASC "

					set rsCats = Server.CreateObject("ADODB.Recordset")
					rsCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

					if rsCats.EOF then
						recCatCnt = ""
					else
						allCatData = rsCats.GetRows(adGetRowsRest)
						recCatCnt = UBound(allCatData,2)
						cCAT_ID = 0
						cCAT_NAME = 1
					end if

					rsCats.close
					set rsCats = nothing

					tmpStrUserList  = ""

					strSql = "SELECT GROUP_CATID "
					strSql = strSql & " FROM " & strTablePrefix & "GROUPS "
					strSql = strSql & " WHERE GROUP_ID = " & cLng("0" & Request.Form("GroupID"))

					set rsGroupCats = Server.CreateObject("ADODB.Recordset")
					rsGroupCats.open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

					if rsGroupCats.EOF then
						recGroupCatCnt = ""
					else
						allGroupCatData = rsGroupCats.GetRows(adGetRowsRest)
						recGroupCatCnt = UBound(allGroupCatData,2)
						gGROUP_CATID = 0
					end if

					rsGroupCats.close
					set rsGroupCats = nothing

					if recGroupCatCnt <> "" then
						for iGroupCats = 0 to recGroupCatCnt
							GroupCatID = allGroupCatData(gGROUP_CATID,iGroupCats)
							if tmpStrUserList = "" then
								tmpStrUserList = GroupCatID
							else
								tmpStrUserList = tmpStrUserList & "," & GroupCatID
							end if
						next
					end if
					SelectSize = 6
					Response.Write "<td>" & strLE & _
						"<table class=""tnb"">" & strLE & _
						"<tr>" & strLE & _
						"<td class=""c""><b>Available</b><br>" & strLE & _
						"<select name=""GroupCatCombo"" size=""" & SelectSize & """ multiple onDblClick=""moveSelectedOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & strLE
					'## Pick from list
					if recCatCnt <> "" then
						for iCat = 0 to recCatCnt
							CategoryCatID = allCatData(cCAT_ID,iCat)
							CategoryCatName = allCatData(cCAT_NAME,iCat)
							if not(Instr("," & tmpStrUserList & "," , "," & CategoryCatID & ",") > 0) then
								Response.Write 	"<option value=""" & CategoryCatID & """>" & ChkString(CategoryCatName,"display") & "</option>" & strLE
							end if
						next
					end if
					Response.Write "</select>" & strLE & _
						"</td>" & strLE & _
						"<td class=""vam c"" width=""15""><br>" & strLE & _
						"<a href=""javascript:moveAllOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemAll,"","class=""vam""") & "</a>" & strLE & _
						"<a href=""javascript:moveSelectedOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & getCurrentIcon(strIconPrivateRemove,"","class=""vam""") & "</a>" & strLE & _
						"<a href=""javascript:moveSelectedOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAdd,"","class=""vam""") & "</a>" & strLE & _
						"<a href=""javascript:moveAllOptions(document.Edit.GroupCatCombo, document.Edit.GroupCat, true, '')"">" & getCurrentIcon(strIconPrivateAddAll,"","class=""vam""") & "</a>" & strLE & _
						"</td>" & strLE & _
						"<td class=""c""><b>Selected</b><br>" & strLE & _
						"<select name=""GroupCat"" size=""" & SelectSize & """ multiple tabindex=""15"" onDblClick=""moveSelectedOptions(document.Edit.GroupCat, document.Edit.GroupCatCombo, true, '')"">" & strLE
					if recGroupCatCnt <> "" then
						for iGroupCats = 0 to recGroupCatCnt
							GroupCatID = allGroupCatData(gGROUP_CATID,iGroupCats)
							if GroupCatID <> "" then
								Response.Write 	"<option value=""" & GroupCatID & """>" & ChkString(getCategoryName(GroupCatID),"display") & "</option>" & strLE
							end if
						next
					end if
					Response.Write "</select>" & strLE & _
						"</td>" & strLE & _
						"</tr>" & strLE & _
						"</table>" & strLE & _
						"</td>" & strLE & _
						"</tr>" & strLE & _
						"<tr>" & strLE & _
						"<td class=""nw c"" colspan=""2""><input class=""button"" value=""Submit"" type=""submit"" tabindex=""5"" onclick=""selectAllOptions(document.Edit.GroupCat);"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset"" tabindex=""6""></td>" & strLE & _
						"</tr>" & strLE & _
						"</table>" & strLE & _
						"</form>" & strLE & _
						"<p class=""c""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></p>" & strLE
				else
					Response.Write "<p class=""c"">Invalid Group ID</span></p>" & strLE & _
						"<br><p class=""c""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></span></p>" & strLE
				end if
			else
				Response.Write "<p class=""c"">Invalid Group ID</span></p>" & strLE & _
					"<br><p class=""c""><a href=""JavaScript:history.go(-1)"">Go back to correct the problem.</a></span></p>" & strLE
			end if
		else
			'## Forum_SQL
			strSql = "SELECT GROUP_ID, GROUP_NAME "
			strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES "
			strSql = strSql & " WHERE GROUP_ID <> 1 "
			strSql = strSql & " ORDER BY GROUP_NAME ASC "

			Set rsgroups = Server.CreateObject("ADODB.Recordset")
			rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

			If rsgroups.EOF then
				recGroupCount = ""
			Else
				allGroupData = rsgroups.GetRows(adGetRowsRest)
				recGroupCount = UBound(allGroupData, 2)
			End if

			rsgroups.Close
			Set rsgroups = Nothing

			Response.Write "<form action=""admin_config_groupcats.asp?method=Edit"" method=""post"" id=""Add"" name=""Add"">" & strLE & _
				"<input type=""hidden"" name=""Method_Type"" value=""Edit_Category"">" & strLE & _
				"<table class=""admin"">" & strLE & _
				"<tr>" & strLE & _
				"<th colspan=""2""><b>Edit Group Categories</b></th>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw r""><b>Choose Group To Edit</b>&nbsp;</td>" & strLE & _
				"<td class=""l"">" & strLE & _
				"<select name=""GroupID"" size=""1"">" & strLE
			if recGroupCount <> "" then
				for iGroup = 0 to recGroupCount
					if allGroupData(0, iGroup) = 2 then
						Response.Write "<option label=""" & chkString(allGroupData(1, iGroup),"display") & """ value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0, iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & strLE
						exit for
					end if
				next
				first = 0
				for iGroup = 0 to recGroupCount
					if allGroupData(0, iGroup) <> 2 then
						if first = 0 then
							Response.Write "<option value="""">----------------------------</option>" & strLE
							first = 1
						end if
						Response.Write "<option value=""" & allGroupData(0, iGroup) & """" & chkSelect(cLng(group),cLng(allGroupData(0, iGroup))) & ">" & chkString(allGroupData(1, iGroup),"display") & "</option>" & strLE
					end if
				next
			end if
			Response.Write "</select>" & strLE & _
				"</td>" & strLE & _
				"</tr>" & strLE & _
				"<tr>" & strLE & _
				"<td class=""nw c"" colspan=""2""><input class=""button"" value=""  Edit  "" type=""submit"">&nbsp;<input name=""Reset"" type=""reset"" value=""Reset""></td>" & strLE & _
				"</tr>" & strLE & _
				"</table>" & strLE & _
				"</form>" & strLE & _
				"<p class=""c""><a href=""admin_config_groupcats.asp"">Back To Group Categories Configuration</a></p>" & strLE
		end if
	Case Else
		Response.Write "<table class=""admin"">" & strLE & _
			"<tr>" & strLE & _
			"<th><b>Group Categories Configuration</b></th>" & strLE & _
			"</tr>" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat""><ul>" & strLE & _
			"<li class=""smt""><a href=""admin_config_groupcats.asp?method=Add"">Create A New Category Group</a></li>" & strLE & _
			"<li class=""smt""><a href=""admin_config_groupcats.asp?method=Delete"">Delete A Category Group</a></li>" & strLE & _
			"<li class=""smt""><a href=""admin_config_groupcats.asp?method=Edit"">Edit an Existing Category Group</a></li>" & strLE & _
			"</ul></td>" & strLE & _
			"</tr>" & strLE & _
			"</table>" & strLE
End Select
WriteFooter
Response.End

sub newGroupCategories(fGroupID)
	if Request.Form("GroupCat") = "" then
		exit Sub
	end if
	Cats = split(Request.Form("GroupCat"),",")
	for count = Lbound(Cats) to Ubound(Cats)
		strSql = "INSERT INTO " & strTablePrefix & "GROUPS ("
		strSql = strSql & " GROUP_ID, GROUP_CATID) VALUES ( "& fGroupID & ", " & Cats(count) & ")"
		my_conn.execute (strSql),,adCmdText + adExecuteNoRecords
	next
end sub

sub updateGroupCategories(fGroupID)
	my_Conn.execute ("DELETE FROM " & strTablePrefix & "GROUPS WHERE GROUP_ID = " & fGroupId),,adCmdText + adExecuteNoRecords
	newGroupCategories(fGroupID)
end sub

Function getCategoryName(fCat_ID)
	set rsCatName = my_Conn.execute("SELECT CAT_NAME FROM " & strTablePrefix & "CATEGORY WHERE CAT_ID = " & fCat_ID)
	getCategoryName = rsCatName("CAT_NAME")
	set rsCatName = nothing
end function
%>
