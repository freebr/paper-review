<!--#include file="pub.asp"-->
<%
Response.Expires=-1
Sub Connect(conn)
	Dim connstr
	connstr=getConnectionString("ThesisReviewSys")
	Set conn=Server.CreateObject("ADODB.Connection")
  conn.CommandTimeout=300
  conn.CursorLocation=adUseClient
	conn.Open connstr
End Sub
Sub ConnectOriginDb(conn)
	Dim connstr
	connstr=getConnectionString("SCUT_MD")
	Set conn=Server.CreateObject("ADODB.Connection")
  conn.CommandTimeout=300
  conn.CursorLocation=adUseClient
	conn.Open connstr
End Sub
Function getConnectionString(initDbName)
	Dim ret
	ret="Provider=SQLNCLI10;Persist Security Info=True;User ID=sa;Password=cnsba2016.net;Initial Catalog="&initDbName
	ret=ret&";Data Source=116.57.68.162,14033;Pooling=true;Max Pool Size=512;Min Pool Size=50;Connection Lifetime=999;"
	getConnectionString=ret
End Function
Function CmdParam(name,ptype,dir,size,value)
	' 构造命令参数对象
	Dim cmd
	Set cmd=Server.CreateObject("ADODB.Command")
	Set CmdParam=cmd.CreateParameter(name,ptype,dir,size,value)
	Set cmd=Nothing
End Function
Function ExecQuery(conn,sql,params,countAffected)
	' 执行查询或存储过程
	Dim cmd,rs
	Set cmd=Server.CreateObject("ADODB.Command")
	If IsEmpty(conn) Then Connect conn
	cmd.ActiveConnection=conn
	cmd.CommandText=sql
	If IsArray(params) Then
		For Each param In params
			cmd.Parameters.Append param
		Next
	Else
		cmd.Parameters.Append params
	End If
	Set rs=cmd.Execute()
	countAffected=rs.RecordCount
	Set ExecQuery=rs
End Function
Function ExecNonQuery(conn,sql,params)
	' 执行不返回记录的存储过程
	Dim cmd:Set cmd=Server.CreateObject("ADODB.Command")
	If IsEmpty(conn) Then Connect conn
	cmd.ActiveConnection=conn
	cmd.CommandText=sql
	If IsArray(params) Then
		For Each param In params
			cmd.Parameters.Append param
		Next
	Else
		cmd.Parameters.Append params
	End If
	cmd.Execute(countAffected)
	ExecNonQuery=countAffected
End Function
%>