<!--#include file="inc/global.inc"--><%

If Request.QueryString()="logout" Then
	If Len(Session("Name")) Then
		msg="行政人员["&Session("Name")&"]登出。"
	ElseIf Len(Session("StuName")) Then
		msg="学生["&Session("StuName")&"]登出。"
	ElseIf Len(Session("TName")) Then
		If Session("IsExpert") Then
			msg="评阅专家["&Session("TName")&"]登出。"
		Else
			msg="教师["&Session("TName")&"]登出。"
		End If
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
	If Len(msg) Then WriteLog msg
	Response.Redirect(redirectUrl)
Else
	Dim usertype,conn,sql,rs,count
	usertype=Request.QueryString("usertype")
	ConnectOriginDb conn
	Select Case usertype
	Case "admin"
		sql="SELECT TEACHERID,TEACHERNO,TEACHERNAME,WRITEPRIVILEGETAGSTRING,READPRIVILEGETAGSTRING FROM TEACHER_INFO WHERE TEACHERID=863 AND VALID=0"
		GetRecordSet conn,rs,sql,count
		Session("Id")=rs("TEACHERID").Value
		Session("No")=rs("TEACHERNO").Value
		Session("Name")=rs("TEACHERNAME").Value
		Session("WritePrivileges")=rs("WRITEPRIVILEGETAGSTRING").Value
		Session("ReadPrivileges")=rs("READPRIVILEGETAGSTRING").Value
		url="admin"
		CloseRs rs
		CloseConn conn
	Case "student"
		stuno=Request.QueryString("no")
		If IsEmpty(stuno) Then stuno="201200000000"
		sql="SELECT * FROM VIEW_STUDENT_INFO WHERE STU_NO="&toSqlString(stuno)&" AND VALID=0"
	  GetRecordSetNoLock conn,rs,sql,count
	  If count>0 Then
		  Session("StuId")=rs("STU_ID").Value
		  Session("StuNo")=rs("STU_NO").Value
		  Session("StuName")=rs("STU_NAME").Value
		  Session("WritePrivileges")=rs("WRITEPRIVILEGETAGSTRING").Value
		  Session("ReadPrivileges")=rs("READPRIVILEGETAGSTRING").Value
		  If rs("TEACH_TYPEID").Value=6 Then
				If InStr(LCase(rs("CLASS_NAME").Value),"mpacc") Then
					Session("StuType")=9
				Else
					Session("StuType")=6
				End If
			Else
				Session("StuType")=rs("TEACH_TYPEID").Value
			End If
		  Session("StuClass")=rs("CLASS_ID").Value
			url="student"
		End If
		CloseRs rs
		CloseConn conn
	Case "tutor"
		username=Request.QueryString("name")
		If IsEmpty(username) Then username="daoshi"
		sql="SELECT TEACHERID,TEACHERNO,TEACHERNAME,WRITEPRIVILEGETAGSTRING,READPRIVILEGETAGSTRING FROM TEACHER_INFO WHERE TEACHERNO='"&username&"' AND VALID=0"
		GetRecordSet conn,rs,sql,count
		If count>0 Then
			Session("TId")=rs("TEACHERID").Value
			Session("TNo")=rs("TEACHERNO").Value
			Session("TName")=rs("TEACHERNAME").Value
			Session("TWritePrivileges")=rs("WRITEPRIVILEGETAGSTRING").Value
			Session("TReadPrivileges")=rs("READPRIVILEGETAGSTRING").Value
			url="tutor"
		End If
		CloseRs rs
		CloseConn conn
	Case "expert"
		username=Request.QueryString("name")
		If IsEmpty(username) Then username="zhuanjia1"
		sql="SELECT TEACHERID,TEACHERNO,TEACHERNAME,WRITEPRIVILEGETAGSTRING,READPRIVILEGETAGSTRING FROM TEACHER_INFO WHERE TEACHERNO='"&username&"' AND VALID=0"
		GetRecordSet conn,rs,sql,count
   	If count>0 Then
		Session("TId")=rs("TEACHERID").Value
		Session("TNo")=rs("TEACHERNO").Value
	    Session("TName")=rs("TEACHERNAME").Value
	    Session("TWritePrivileges")=rs("WRITEPRIVILEGETAGSTRING").Value
	    Session("TReadPrivileges")=rs("READPRIVILEGETAGSTRING").Value
	    Session("IsExpert")=True
		url="expert"
	  End If
   	CloseRs rs
		CloseConn conn
	Case Else
		Response.End()
	End Select
	
	If Len(url)=0 Then
		Response.Redirect("error.asp?user-not-exists")
	End If
	Response.Redirect url
End If
%>