<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")%><%
status=Request.Form("status")
Select Case status
Case "open"
	setSystemStatus "open"
Case "closed"
	setSystemStatus "closed"
Case "debug"
	setSystemStatus "debug"
End Select
CloseConn conn
%>{ status: "ok" }