<!--#include file="adovbs.inc"-->
<!--#include file="config.inc"-->
<script language="jscript" runat="server">
	function ExecQuery(conn,sql) {
		// 执行查询或存储过程
		conn = conn || ConnectDb();
		var cmd=new ActiveXObject("ADODB.Command");
		cmd.ActiveConnection=conn;
		cmd.CommandText=sql;
		for (var i=2;i<arguments.length;++i) {
			if (arguments[i]) cmd.Parameters.Append(arguments[i]);
		}
		var rs=cmd.Execute();
		var dict=CreateDictionary();
		dict.Add("rs", rs);
		dict.Add("count", rs.RecordCount);
		return dict;
	}
	function ExecNonQuery(conn,sql) {
		// 执行不返回记录的存储过程
		conn = conn || ConnectDb();
		var countAffected=0;
		var cmd=new ActiveXObject("ADODB.Command");
		cmd.ActiveConnection=conn;
		cmd.CommandText=sql;
		for (var i=2;i<arguments.length;++i) {
			if (arguments[i]) cmd.Parameters.Append(arguments[i]);
		}
		cmd.Execute(countAffected);
		return countAffected;
	}
</script><%
' 建立本系统数据库连接
Function ConnectDb(conn)
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.CommandTimeout=300
	conn.CursorLocation=adUseClient
	conn.Open getConnectionString(uriDatabaseServer, "PaperReviewSystem")
	Set ConnectDb = conn
End Function
' 建立教务系统数据库连接
Function ConnectJWDb(conn)
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.CommandTimeout=300
	conn.CursorLocation=adUseClient
	conn.Open getConnectionString(uriJWDatabaseServer, "SCUT_MD")
	Set ConnectJWDb = conn
End Function
Function getConnectionString(uriServer,initDbName)
	Dim ret
	ret="Provider=SQLNCLI10;Persist Security Info=True;User ID=PaperReviewSystem;Password=HgggLwpy@87114057;Initial Catalog="&initDbName
	ret=ret&Format(";Data Source={0};Pooling=true;Max Pool Size=512;Min Pool Size=50;Connection Lifetime=999;",uriServer)
	getConnectionString=ret
End Function
' 构造命令参数对象
Function CmdParam(name,ptype,size,value)
	Dim cmd
	Set cmd=Server.CreateObject("ADODB.Command")
	Set CmdParam=cmd.CreateParameter(name,ptype,adParamInput,size,value)
	Set cmd=Nothing
End Function

Sub GetRecordSet(conn,rs,sqlStr,count)
	Set rs=Server.CreateObject("ADODB.RECORDSET")
	If IsEmpty(conn) Then ConnectDb conn
	rs.ActiveConnection=conn
	rs.Source=sqlStr
	rs.Open , ,AdOpenKeyset,AdLockOptimistic
	count=rs.RecordCount
End Sub
'========================

Sub GetRecordSetNoLock(conn,rsNoLock,sqlStr,count)
	Set rsNoLock=Server.CreateObject("ADODB.RECORDSET")
	If IsEmpty(conn) Then ConnectDb conn
	rsNoLock.activeConnection=conn
	rsNoLock.source=sqlStr
	rsNoLock.Open , ,AdOpenKeyset,AdLockReadOnly
	count=rsNoLock.RecordCount
End Sub

'=======================
Sub GetMenuListPubTerm(table,FieldID,FieldName,fieldValue,TermStr)
	Set rsMenu=Server.CreateObject("ADODB.RECORDSET")
	If IsEmpty(conn) Then ConnectDb conn
	rsMenu.activeConnection=conn
	If FieldID="" Then
		rsMenu.source="SELECT "
	Else
		rsMenu.source="SELECT DISTINCT "&FieldID&","
	End If
	If TermStr="" Then
		rsMenu.source=rsMenu.source&fieldName&" FROM "&table&" WHERE VALID=1 "
	Else
		rsMenu.source=rsMenu.source&fieldName&" FROM "&table&" WHERE VALID=1 "&TermStr
	End If
	rsMenu.Open , ,AdOpenKeyset
	While Not rsMenu.EOF
		If rsMenu(fieldName)<>"" Then
			Response.write "<OPTION VALUE='"&rsMenu(FieldID)&"'"
			If Cstr(rsMenu(fieldID))=Cstr(fieldValue) Then Response.write " SELECTED "
			Response.write ">"&rsMenu(fieldName)&"</OPTION>"&vbcrlf
		End If
		rsMenu.MoveNext()
	Wend
	Set rsMenu=Nothing
End Sub

Sub CloseConn(conn)
	If Not IsEmpty(conn) Then Set conn=Nothing
End Sub
Sub CloseRs(rs)
	If Not IsEmpty(rs) Then Set rs=Nothing
End Sub
%>