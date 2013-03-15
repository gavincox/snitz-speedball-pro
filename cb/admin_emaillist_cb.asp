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
	"function ChangePage(fnum){" & strLE & _
	"if (fnum == 1) {" & strLE & _
	"document.PageNum1.submit();" & strLE & _
	"}" & strLE & _
	"else {" & strLE & _
	"document.PageNum2.submit();" & strLE & _
	"}" & strLE & _
	"}" & strLE & _
	"</script>" & strLE
mypage = trim(chkString(request("whichpage"),"SQLString"))
if ((mypage = "") or (IsNumeric(mypage) = FALSE)) then mypage = 1
mypage = cLng(mypage)
'## Forum_SQL - Find all records with the search criteria in them
strSql = "SELECT M_NAME, M_EMAIL, M_POSTS "
strSql2 = " FROM " & strMemberTablePrefix & "MEMBERS "
strSql3 = " WHERE M_STATUS = " & 1
strSql4 = " ORDER BY MEMBER_ID ASC "
if strDBType = "mysql" then 'MySql specific code
	if mypage > 1 then
		OffSet = cLng((mypage - 1) * strPageSize)
		strSql5 = " LIMIT " & OffSet & ", " & strPageSize & " "
	end if
	'## Forum_SQL - Get the total pagecount
	strSql1 = "SELECT COUNT(MEMBER_ID) AS PAGECOUNT "
	set rsCount = my_Conn.Execute(strSql1 & strSql2 & strSql3)
	iPageTotal = rsCount(0).value
	rsCount.close
	set rsCount = nothing
	if iPageTotal > 0 then
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
		rs.open strSql & strSql2 & strSql3 & strSql4 & strSql5, my_Conn, adOpenForwardOnly, adLockReadOnly, adCmdText
			arrMemberData = rs.GetRows(intGetRows)
			iMemberCount = UBound(arrMemberData, 2)
		rs.close
		set rs = nothing
	else
		iTopicCount = ""
	end if
else 'end MySql specific code
	set rs = Server.CreateObject("ADODB.Recordset")
	rs.cachesize = strPageSize
	rs.open strSql & strSql2 & strSql3 & strSql4, my_Conn, adOpenStatic
		If not (rs.EOF or rs.BOF) then
			rs.movefirst
			rs.pagesize = strPageSize
			rs.absolutepage = mypage '**
			maxpages = cLng(rs.pagecount)
			arrMemberData = rs.GetRows(strPageSize)
			iMemberCount = UBound(arrMemberData, 2)
		else
			iMemberCount = ""
		end if
	rs.Close
	set rs = nothing
end if
sub DropDownPaging(fnum)
	if maxpages > 1 then
		if mypage = "" then
			pge = 1
		else
			pge = mypage
		end if
		scriptname = request.servervariables("script_name")
		Response.Write "<form name=""PageNum" & fnum & """ action=""admin_emaillist.asp"">" & strLE
		Response.Write "<td><span class=""dff dfs"">" & strLE
		if fnum = 1 then
			Response.Write("<b>Page: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
		else
			Response.Write("<b>There are " & maxpages & " Pages of Members: </b><select name=""whichpage"" size=""1"" onchange=""ChangePage(" & fnum & ");"">" & strLE)
		end if
		for counter = 1 to maxpages
			if counter <> cLng(pge) then
				Response.Write "<option value=""" & counter &  """>" & counter & "</option>" & strLE
			else
				Response.Write "<option selected value=""" & counter &  """>" & counter & "</option>" & strLE
			end if
		next
		if fnum = 1 then
			Response.Write("</select><b> of " & maxPages & "</b>" & strLE)
		else
			Response.Write("</select>" & strLE)
		end if
		Response.Write("</span></td>" & strLE)
		Response.Write("</form>" & strLE)
	end if
end sub
%>
