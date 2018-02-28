<!-- #include File="inc/pub.asp" --><%
If Len(Session("name")) Then
	msg="行政人员["&Session("name")&"]登出."
End If
If Session("LoginFromCAS") Then
	Dim ids,gotoUrl
	Set ids=newLoginComp()
	gotoUrl="http://"&Request.ServerVariables("SERVER_NAME")&"/"
	redirectUrl=ids.getLogoutUrl()&"?goto="&Server.URLEncode(gotoUrl)
	msg=msg&"(通过统一认证系统)"
Else
	redirectUrl="/"
End If
Session.Abandon()
WriteLog msg
Response.Redirect(redirectUrl)
%>