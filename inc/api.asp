<%
If IsEmpty(Session("Id")) And IsEmpty(Session("TId")) And IsEmpty(Session("StuId")) Then
    Response.Redirect("../error.asp?timeout")
End If

Function ensureArgument(args, params, dict)
    If Not IsArray(args) Then args = Array(args)
    Dim i
    For i=0 To UBound(args)
        If IsEmpty(Request(args(i))) Then
            dict.Add "status", "error"
            dict.Add "msg", Format("Parameter missing: '{0}'", Array(args(i)))
            Call (new JSONWriter)(Response, dict)
            Response.End()
            ensureArgument=False
            Exit Function
        End If
        params(args(i)) = Request(args(i))
    Next
    ensureArgument=True
End Function

%>