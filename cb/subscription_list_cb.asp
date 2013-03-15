<%
sub Go_Result
	' Go_Result - Closes connections, displays footer, etc
	if iSubCount = 0 then
		Response.Write "<p class=""c""><b>No Subscriptions found</b></p>" & strLE & _
			"<p class=""c""><a href=""JavaScript:history.go(-1)"">Go Back To Forum</a></p>" & strLE
	end if
	set rs = nothing ' -- Close all connections
	Response.Write "</table>" & strLE
	Call WriteFooter
	Response.End
end sub

Function GetSubLevel(CurrLevel)
	Dim Textout : Textout = ""
	if CurrLevel = 0 then
		Textout = " (No Subscriptions allowed)"
	else
		Textout = " (Subscription level set to "
		Select Case CurrLevel
			Case 1 : Textout = Textout & "Category)"
			Case 2 : Textout = Textout & "Forum)"
			Case 3 : Textout = Textout & "Topic)"
			Case else : Textout = "(??)"
		End Select
	End if
	GetSubLevel = "<span class=""ffs"">" & Textout & "</span>"
End Function

Function GetFSubLevel(CurrLevel)
	Dim Textout : Textout = ""
	if CurrLevel = 0 then
		Textout = " (No Subscriptions allowed)"
	else
		Textout = " (Subscription level set to "
		Select Case CurrLevel
			Case 1 : Textout = Textout & "Forum)"
			Case 2 : Textout = Textout & "Topic)"
			Case else : Textout = "(??)"
		End Select
	End if
	GetFSubLevel = "<span class=""ffs"">" & Textout & "</span>"
End Function
%>