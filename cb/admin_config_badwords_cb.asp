<%
sub DropDownPaging()
    if maxpages > 1 then
        if mypage = "" then pge = 1 else pge = mypage
        Response.Write "<td class=""ccc vam""><span class=""dff ffs cfc"">" & strLE & _
            "<b>Page</b>&nbsp;<select style=""font-size:9px"" name=""whichpage"" size=""1"" onchange=""jumpToPage(this)"">" & strLE
        for counter = 1 to maxpages
            ref = "admin_config_badwords.asp?whichpage=" & counter
            if counter <> cLng(pge) then
                Response.Write "<option value=""" & ref & """>" & counter & "</option>" & strLE
            else
                Response.Write "<option value=""" & ref & """ selected>" & counter & "</option>" & strLE
            end if
        next
        Response.Write "</select>&nbsp;<b>of " & maxpages & "</b></span></td>" & strLE
    end if
end sub
function chkBString(fString,fField_Type) '## Types - SQLString
    if fString = "" then fString = " "
    Select Case fField_Type
        Case "SQLString"
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
            fString    = HTMLBEncode(fString)
            chkBString = fString
            exit function
    End Select
    chkBString = fString
end function
function HTMLBEncode(fString)
    if fString = "" or IsNull(fString) then fString = " "
    fString     = replace(fString, ">", "&gt;")
    fString     = replace(fString, "<", "&lt;")
    HTMLBEncode = fString
end function
%>