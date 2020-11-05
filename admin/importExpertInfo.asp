<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

tableUploadPath = Server.MapPath(uploadBasePath(usertypeAdmin, "expert_info"))
ensurePathExists tableUploadPath

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入专家信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body>
<center><font size=4><b>导入专家信息</b><br>
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<input type="file" name="tableFile" size="100" /></p>
<p><input type="submit" name="btnupload" value="上 传" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$(document).ready(function(){
		$('form').submit(function() {
			var valid=checkIfExcel(this.tableFile);
			if(valid) {
				$(':submit').val("正在提交，请稍候...").attr('disabled',true);
			}
			return valid;
		});
		$(':submit').attr('disabled',false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim Upload,File
	
	Set Upload=New ExtendedRequest
	Set file=Upload.File("tableFile")
	
	file_ext=LCase(file.FileExt)
	If file_ext <> "xls" And file_ext <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errMsg = "所选择的不是 Excel 文件！"
	Else
		destFile = timestamp()&"."&file_ext
		destPath = resolvePath(tableUploadPath,destFile)
		file.SaveAs destPath
	End If
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入专家信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="filename" value="<%=destFile%>" />
<p>文件上传成功，正在导入专家信息...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim fieldValue(6)
		Dim sql,conn,connOrigin,count,rsExp,rsTea,rsTmp
		Dim numNewTeacher,numNewExTeacher,numUpdTeacher
		Dim bIsUpdated,bIsInschool
		Dim arrRet(1)
		Dim py,username
		Dim s,i
		numUpdTeacher=0
		numNewTeacher=0
		numNewExTeacher=0
		Set py=New PinyinQuery
		ConnectDb conn
		ConnectJWDb connOrigin
		sql="SELECT * FROM Experts"
		GetRecordSet conn,rsExp,sql,count
		sql="SELECT * FROM TEACHER_INFO"
		GetRecordSet connOrigin,rsTea,sql,count
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			bIsUpdated=False
			' 是否校外专家
			If IsNull(rs(6)) Or rs(6)="否" Then
				bIsInschool=True
			Else
				bIsInschool=False
			End If
			
			' 姓名
			s=Trim(rs(0))
			If bIsInschool Then i=getTeacherIdByName(s)
			If bIsInschool And i=-1 Then
				bError=True
				errMsg=errMsg&""""&s&"""按校内教师导入，但该教师不存在。"&vbNewLine
			Else
				fieldValue(0)=s
				' 职称
				fieldValue(1)=rs(1)
				' 学科专长
				fieldValue(2)=rs(2)
				' 单位（住址）
				fieldValue(3)=rs(3)
				' 联系方式
				fieldValue(4)=rs(4)
				' 邮箱
				fieldValue(5)=rs(5)
				' 备注
				fieldValue(6)=rs(6)
				
				sql="EXPERT_NAME='"&fieldValue(0)&"' AND INSCHOOL="&Abs(Int(bIsInschool))
				rsExp.Filter=sql
				If rsExp.EOF Then	' 添加记录
					rsExp.AddNew()
					If Not bIsInschool Then	' 添加校外专家到教师信息表
						rsTea.AddNew()
						' 生成登录用户名，并保证不与已有用户名相同
						username=LCase(Replace(py.getNamePinyinOf(fieldValue(0))," ",""))
						s=username
						i=0
						Do
							If i>0 Then	s=username&i
							sql="SELECT TEACHERID FROM TEACHER_INFO WHERE TEACHERNO="&toSqlString(s)
							GetRecordSetNoLock connOrigin,rsTmp,sql,count
							CloseRs rsTmp
							i=i+1
						Loop While count>0
						rsTea("TEACHERNO")=s
						rsTea("TEACHERNAME")=fieldValue(0)
						rsTea("USER_PASSWORD")="expert@12345" ' generatePassword()
						rsTea("IFTEACHER")=3
						rsTea("Office_Address")=fieldValue(3)
						rsTea("MOBILE")=fieldValue(4)
						rsTea("EMAIL")=fieldValue(5)
						rsTea("PRO_DUTYID")=18	' 职称为其他
						rsTea.Update()
						numNewExTeacher=numNewExTeacher+1
					End If
					numNewTeacher=numNewTeacher+1
				Else							' 更新记录
					If Not bIsInschool Then	' 更新校外专家在教师信息表中的记录
						rsTea.Find("TEACHERID="&rsExp("TEACHER_ID"))
						If Not rsTea.EOF Then
							rsTea("USER_PASSWORD")="12345678" ' generatePassword()
							rsTea("Office_Address")=fieldValue(3)
							rsTea("MOBILE")=fieldValue(4)
							rsTea("EMAIL")=fieldValue(5)
							rsTea("VALID")=0
							rsTea.Update()
						End If
					End If
					numUpdTeacher=numUpdTeacher+1
				End If
				rsExp("EXPERT_NAME")=fieldValue(0)
				rsExp("PRO_DUTY_NAME")=fieldValue(1)
				rsExp("EXPERTISE")=fieldValue(2)
				rsExp("WORKPLACE")=fieldValue(3)
				rsExp("MOBILE")=fieldValue(4)
				rsExp("EMAIL")=fieldValue(5)
				rsExp("MEMO")=fieldValue(6)
				rsExp("INSCHOOL")=Abs(bIsInschool)
				rsExp.Update()
			End If
			rs.MoveNext()
		Loop
		CloseRs rsTea
		CloseRs rsExp
		' 调用绑定教师ID的存储过程
		conn.Execute "EXEC spUpdateExpertId"
		' 添加专家权限，使其可以进入评阅系统
		conn.Execute "EXEC spConfigExpertPrivilege"
		CloseConn connOrigin
		CloseConn conn
		Set py=Nothing
		arrRet(0)=numNewTeacher
		arrRet(1)=numUpdTeacher
		addData=arrRet
	End Function
	
	Dim bError,errMsg,arr
	
	filename=Request.Form("filename")
	filepath=resolvePath(tableUploadPath,filename)
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&filepath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
	connExcel.Open connstring
	
	Set rs=connExcel.OpenSchema(adSchemaTables)
	Do While Not rs.EOF
		If rs("TABLE_TYPE")="TABLE" Then
			table_name=rs("TABLE_NAME")
			If InStr("Sheet1$",table_name) Then Exit Do
		End If
		rs.MoveNext()
	Loop
	
	sql="SELECT * FROM ["&table_name&"A2:H]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	arr=addData()
	CloseRs rs
	CloseConn connExcel
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，<%=arr(0)%>条记录已导入，<%=arr(1)%>条记录已更新。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=arr(0)%>条记录已导入，<%=arr(1)%>条记录已更新。");
<%End If
%>location.href="expertList.asp";
</script><%
End Select
%>