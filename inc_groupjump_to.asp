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
if strGroupCategories = "1" then
	strOK = ""
	Response.Write "<script type=""text/javascript"">" & strLE & _
		"<!--" & strLE & _
		"function jumpTo(s) {if (s.selectedIndex != 0) location.href = s.options[s.selectedIndex].value;return 1;}" & strLE & _
		"// -->" & strLE & _
		"</script>" & strLE

	' where we are?
	strPathInfo = Request.ServerVariables("Path_Info")
	if lcase(Right(strPathInfo, 10)) = "active.asp" Then
		strOK = "OK"
		strLinkTo = "active.asp"
	elseif lcase(Right(strPathInfo, 11)) = "default.asp" then
		strOK = "OK"
		strLinkTo = "default.asp"
	else
		strOK = ""
	end if

	if StrOK="OK" then
		Response.Write "<form name=""GroupStuff"">" & strLE & _
			"<tr>" & strLE & _
			"<td class=""vat"">" & strLE & _
			"<span class=""dff dfs""><b>Change Category Group<br></b></span>" & strLE & _
			"<select name=""SelectMenu"" size=""0"" onchange=""if(this.options[this.selectedIndex].value != '' ){ jumpTo(this) }"">" & strLE & _
			"<option value="""">Select Other Categories Here!</option>" & strLE

		'## Get all Forum Groups From DB
		strSql = "SELECT GROUP_ID, GROUP_NAME"
		strSql = strSql & " FROM " & strTablePrefix & "GROUP_NAMES"
		strSql = strSql & " ORDER BY GROUP_NAME ASC;"

		set rsgroups = Server.CreateObject("ADODB.Recordset")
		rsgroups.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText

		if rsgroups.EOF then
			recGroupCatCount = ""
	 	else
	 		allGroupCatData = rsgroups.GetRows(adGetRowsRest)
	 		recGroupCatCount = UBound(allGroupCatData, 2)
	 	end if

	 	rsgroups.Close
	 	set rsgroups = nothing

		if recGroupCatCount <> "" then
			gGROUP_ID = 0
			gGROUP_NAME = 1

			for iGroup = 0 to recGroupCatCount
				GroupID = allGroupCatData(gGROUP_ID,iGroup)
				GroupName = allGroupCatData(gGROUP_NAME,iGroup)

				if GroupID = 1 or GroupID = 2 then
					Response.Write "<option value=""" & strLinkTo & "?Group=" & GroupID & """" & chkSelect(cLng(Group), cLng(GroupID)) & ">" & GroupName & "</option>" & strLE
				end if
			next
			first = 0
			for iGroup = 0 to recGroupCatCount
				GroupID = allGroupCatData(gGROUP_ID,iGroup)
				GroupName = allGroupCatData(gGROUP_NAME,iGroup)

				if GroupID = 1 OR GroupID = 2 then
					' do nothing
				else
					if first = 0 then
						Response.Write "<option value="""">----------------------------</option>" & strLE
						first = 1
					end if
					Response.Write "<option value=""" & strLinkTo & "?Group=" & GroupID & """" & chkSelect(cLng(Group), cLng(GroupID)) & ">" & GroupName & "</option>" & strLE
				end if
			next
		end if
		Response.Write "</select>" & strLE & _
			"<span class=""dff dfs""><a href=""default_group.asp""><acronym title=""Group Categories links and Information"">Group Category Menu</acronym></a></span></td>" & strLE & _
			"</tr>" & strLE & _
			"</form>" & strLE & _
			"<tr>" & strLE & _
			"<td><span style=""font-size: 6px;""><br></span></td>" & strLE & _
			"</tr>" & strLE
	end if
end if
%>
