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
if Session(strCookieURL & "Approval") <> "15916941253" then
	scriptname = split(request.servervariables("SCRIPT_NAME"),"/")
	Response.Redirect "admin_login.asp?target=" & scriptname(ubound(scriptname))
end if
Response.Write "<script type=""text/javascript"">" & strLE & _
	"<!-- " & vbNewLine & _
	"function ChangePage(fnum){" & strLE & _
	"if (fnum == 1) {" & strLE & _
	"document.PageNum1.submit();" & strLE & _
	"}" & strLE & _
	"else {" & strLE & _
	"document.PageNum2.submit();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function appr_all(){" & strLE & _
	"var where_to= confirm(""Do you really want to Approve all Pending Members?"");" & strLE & _
	"if (where_to== true) {" & strLE & _
	"window.location=""admin_accounts_pending.asp?id=-1&action=approve"";" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function appr_selected(){" & strLE & _
	"var where_to= confirm(""Do you really want to Approve the Selected Pending Members?"");" & strLE & _
	"if (where_to== true) {" & strLE & _
	"document.delMembers.action.value = 'approve';" & strLE & _
	"document.delMembers.submit();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function del_all(){" & strLE & _
	"var where_to= confirm(""Do you really want to Delete all Pending Members?"");" & strLE & _
	"if (where_to== true) {" & strLE & _
	"window.location=""admin_accounts_pending.asp?id=-1&action=delete"";" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function del_selected(){" & strLE & _
	"var where_to= confirm(""Do you really want to Delete the Selected Pending Members?"");" & strLE & _
	"if (where_to== true) {" & strLE & _
	"document.delMembers.action.value = 'delete';" & strLE & _
	"document.delMembers.submit();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function Toggle(field)" & strLE & _
	"{" & strLE & _
	"if (field.checked) {" & strLE & _
	"document.delMembers.toggleAll.checked = AllChecked();" & strLE & _
	"}" & strLE & _
	"else {" & strLE & _
	"document.delMembers.toggleAll.checked = false;" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function ToggleAll(field)" & strLE & _
	"{" & strLE & _
	"if (field.checked) {" & strLE & _
	"CheckAll();" & strLE & _
	"}" & strLE & _
	"else {" & strLE & _
	"ClearAll();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"function Check(field)" & strLE & _
	"{" & strLE & _
	"field.checked = true;" & strLE & _
	"}" & strLE & _
	"function Clear(field)" & strLE & _
	"{" & strLE & _
	"field.checked = false;" & strLE & _
	"}" & strLE & _
	"function CheckAll()" & strLE & _
	"{" & strLE & _
	"var dm = document.delMembers;" & strLE & _
	"var len = dm.elements.length;" & strLE & _
	"for (var i = 0; i < len; i++) {" & strLE & _
	"var field = dm.elements[i];" & strLE & _
	"if (field.name == ""id"") {" & strLE & _
	"Check(field);" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"dm.toggleAll.checked = true;" & strLE & _
	"}" & strLE & _
	"function ClearAll()" & strLE & _
	"{" & strLE & _
	"var dm = document.delMembers;" & strLE & _
	"var len = dm.elements.length;" & strLE & _
	"for (var i = 0; i < len; i++) {" & strLE & _
	"var field = dm.elements[i];" & strLE & _
	"if (field.name == ""id"") {" & strLE & _
	"Clear(field);" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"dm.toggleAll.checked = false;" & strLE & _
	"}" & strLE & _
	"function AllChecked()" & strLE & _
	"{" & strLE & _
	"dm = document.delMembers;" & strLE & _
	"len = dm.elements.length;" & strLE & _
	"for(var i = 0 ; i < len ; i++) {" & strLE & _
	"if (dm.elements[i].name == ""id"" && !dm.elements[i].checked) {" & strLE & _
	"return false;" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"return true;" & strLE & _
	"}" & strLE & _
	"//-->" & strLE & _
	"</script>" & strLE
mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)
if mypage > 1 then strPage = "?whichpage=" & mypage
selID = Request.QueryString("id")
strAction = Request.QueryString("action")
if strAction = "approve" then
	if selID = "-1" then
		Call EmailMembers("all")
		'## Forum_SQL - Approve all members
		strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING"
		strSql = strSql & " SET M_APPROVE = " & 1
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		Response.Write "<br><p class=""c""><span class=""dff hfs""><b>Members Approved!</b></span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""5; URL=admin_accounts_pending.asp" & strPage & """>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">All Pending Members have been approved! Their registration e-mails have been sent to them.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""admin_accounts_pending.asp" & strPage & """>Back To Members Pending</span></a></p><br>" & strLE
		Call WriteFooter
		Response.End
	else
		Call EmailMembers("selected")
		aryID = split(selID, ",")
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Approve all members
			strSql = "UPDATE " & strMemberTablePrefix & "MEMBERS_PENDING"
			strSql = strSql & " SET M_APPROVE = " & 1
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		next
		Response.Write "<br><p class=""c""><span class=""dff hfs""><b>Members Approved!</b></span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""5; URL=admin_accounts_pending.asp" & strPage & """>" & strLE & _
			"<p class=""c""><span class=""dff dfs"">Selected Pending Members have been approved! Their registration e-mails have been sent to them.</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""admin_accounts_pending.asp" & strPage & """>Back To Members Pending</span></a></p><br>" & strLE
		Call WriteFooter
		Response.End
	end if
elseif strAction = "delete" then
	if selID = "-1" then
		'## Forum_SQL - Delete the Member
		strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
		strSql = strSql & " WHERE M_STATUS = " & 0
		strSql = strSql & " AND M_LEVEL = " & -1
		my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		Response.Write "<br><p class=""c""><span class=""dff hfs""><b>Members Deleted!</b></span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_accounts_pending.asp" & strPage & """>" & strLE & _
			"<p class=""c""><span class=""dff hfs"">All pending members have been deleted!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""admin_accounts_pending.asp" & strPage & """>Back To Members Pending</span></a></p><br>" & strLE
		Call WriteFooter
		Response.End
	else
		aryID = split(selID, ",")
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Delete the Member
			strSql = "DELETE FROM " & strMemberTablePrefix & "MEMBERS_PENDING "
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)
			my_Conn.Execute (strSql),,adCmdText + adExecuteNoRecords
		next
		Response.Write "<br><p class=""c""><span class=""dff hfs""><b>Members Deleted!</b></span></p>" & strLE & _
			"<meta http-equiv=""Refresh"" content=""2; URL=admin_accounts_pending.asp" & strPage & """>" & strLE & _
			"<p class=""c""><span class=""dff hfs"">Selected members have been deleted!</span></p>" & strLE & _
			"<p class=""c""><span class=""dff dfs""><a href=""admin_accounts_pending.asp" & strPage & """>Back To Members Pending</span></a></p><br>" & strLE
		Call WriteFooter
		Response.End
	end if
end if
'## Forum_SQL - Find all records with the search criteria in them
strSql = "SELECT M_NAME, M_EMAIL, MEMBER_ID, M_DATE, M_IP, M_KEY, M_APPROVE"
strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
strSql3 = " ORDER BY MEMBER_ID ASC"
if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		OffSet = cLng((mypage - 1) * strPageSize)
		strSql4 = " LIMIT " & OffSet & ", " & strPageSize & " "
	end if
	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "
	set rsCount = my_Conn.Execute(strSql1 & strSql2)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing
	If iPageTotal > 0 then
		maxpages = (iPageTotal \ strPageSize )
		if iPageTotal mod strPageSize <> 0 then maxpages = maxpages + 1
		if iPageTotal < (strPageSize + 1) then
			intGetRows = iPageTotal
		elseif (mypage * strPageSize) > iPageTotal then
			intGetRows = strPageSize - ((mypage * strPageSize) - iPageTotal)
		else
			intGetRows = strPageSize
		end if
	else
		iPageTotal = 0
		maxpages   = 0
	end if
	if iPageTotal > 0 then
		set rs = Server.CreateObject("ADODB.Recordset")
		rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		arrMemberData = rs.GetRows(intGetRows)
		iMemberCount = UBound(arrMemberData, 2)
		rs.close
		set rs = nothing
	else
		iMemberCount = ""
	end if
else 'end MySql specific code
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3, my_Conn, adOpenStatic
		if not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize     = strPageSize
			rs.absolutepage = mypage '**
			maxpages        = cLng(rs.pagecount)
			if maxpages >= mypage then
				arrMemberData = rs.GetRows(strPageSize)
				iMemberCount = UBound(arrMemberData, 2)
			else
				iMemberCount = ""
			end if
		else
			iMemberCount = ""
		end if
	rs.Close
	set rs = nothing
end if
sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then pge = 1 else pge = mypage
		scriptname = request.servervariables("script_name")
		Response.Write "<form name=""PageNum" & fnum & """ action=""admin_accounts_pending.asp"">" & strLE
		Response.Write "<td><span class=""dff dfs"">" & strLE
		if fnum = 1 then
			Response.Write "<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE
		else
			Response.Write "<b>There are " & maxpages & " Pages of Pending Members: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		if fnum = 1 then
			Response.Write "</select><b> of " & maxPages & "</b>" & strLE
		else
			Response.Write "</select>" & strLE
		end if
		Response.Write "</span></td>" & strLE
		Response.Write "</form>" & strLE
	end if
end sub
sub EmailMembers(who)
	if who = "all" then
		'## Forum_SQL - Get all pending members
		strSql = "SELECT M_NAME, M_EMAIL, M_KEY, M_APPROVE"
		strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
		strSql = strSql & " ORDER BY MEMBER_ID ASC"
		set rsApprove = Server.CreateObject("ADODB.Recordset")
		rsApprove.Open strSql, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
		if rsApprove.EOF then
			recApproveCount = ""
		else
			allApproveData = rsApprove.GetRows(adGetRowsRest)
			recApproveCount = UBound(allApproveData, 2)
		end if
		rsApprove.Close
		set rsApprove = Nothing
		if recApproveCount <> "" then
			mM_NAME = 0
			mM_EMAIL = 1
			mM_KEY = 2
			mM_APPROVE = 3
			for RowCount = 0 to recApproveCount
				MP_MemberName    = allApproveData(mM_NAME,RowCount)
				MP_MemberEMail   = allApproveData(mM_EMAIL,RowCount)
				MP_MemberKey     = allApproveData(mM_KEY,RowCount)
				MP_MemberApprove = allApproveData(mM_APPROVE,RowCount)
				if MP_MemberApprove = 0 then
					'## E-mails Message to all pending members.
					strRecipientsName = MP_MemberName
					strRecipients     = MP_MemberEMail
					strFrom           = strSender
					strFromName       = strForumTitle
					strsubject        = strForumTitle & " Registration "
					strMessage        = "Hello " & MP_MemberName & strLE & strLE
					strMessage        = strMessage & "You received this message from " & strForumTitle & " because you have registered for a new account which allows you to post new messages and reply to existing ones on the forums at " & strForumURL & strLE & strLE
					if strAuthType = "db" then
						strMessage = strMessage & "Please click on the link below to complete your registration." & strLE & strLE
						strMessage = strMessage & strForumURL & "register.asp?actkey=" & MP_MemberKey & strLE & strLE
					end if
					strMessage = strMessage & "You can change your information at our website by selecting the ""Profile"" link." & strLE & strLE
					strMessage = strMessage & "Happy Posting!"
%>
<!--#INCLUDE VIRTUAL="/inc_mail.asp" -->
<%
				end if
			next
		end if
	elseif who = "selected" then
		aryID = split(selID, ",")
		for i = 0 to ubound(aryID)
			'## Forum_SQL - Get all pending members
			strSql = "SELECT M_NAME, M_EMAIL, M_KEY, M_APPROVE"
			strSql = strSql & " FROM " & strMemberTablePrefix & "MEMBERS_PENDING"
			strSql = strSql & " WHERE MEMBER_ID = " & aryID(i)
			set rsApprove = my_Conn.Execute(strSql)
			if not(rsApprove.EOF) and not(rsApprove.BOF) and rsApprove("M_APPROVE") = 0 then
				'## E-mails Message to all pending members.
				strRecipientsName = rsApprove("M_NAME")
				strRecipients     = rsApprove("M_EMAIL")
				strFrom           = strSender
				strFromName       = strForumTitle
				strsubject        = strForumTitle & " Registration "
				strMessage        = "Hello " & rsApprove("M_NAME") & strLE & strLE
				strMessage        = strMessage & "You received this message from " & strForumTitle & " because you have registered for a new account which allows you to post new messages and reply to existing ones on the forums at " & strForumURL & strLE & strLE
				if strAuthType="db" then
					strMessage = strMessage & "Please click on the link below to complete your registration." & strLE & strLE
					strMessage = strMessage & strForumURL & "register.asp?actkey=" & rsApprove("M_KEY") & strLE & strLE
				end if
				strMessage = strMessage & "You can change your information at our website by selecting the ""Profile"" link." & strLE & strLE
				strMessage = strMessage & "Happy Posting!"
%>
<!--#INCLUDE VIRTUAL="/inc_mail.asp" -->
<%
				rsApprove.movenext
			end if
			rsApprove.Close
			set rsApprove = nothing
		next
	end if
end sub
%>
