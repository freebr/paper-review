<%Response.Charset="utf-8"%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 文件选择页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入自EXCEL文件</b><br>
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" />
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center>
<script type="text/javascript">
	$('#fmUpload').excelFile.onchange=function() {
		var fileName = this.value;
		var fileExt = fileName.substring(fileName.lastIndexOf('.')).toLowerCase();
		if (fileExt != ".xls" && fileExt != ".xlsx") {
			alert("所选文件不是 Excel 文件！");
			this.form.reset();
			return false;
		}
		this.form.submit();
	}
</script></body></html><%
Case 2	' 上传进程

	Dim fso,Upload,File
	
	Set Upload=New ExtendedRequest
	Set file=Upload.File("excelFile")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 检查上传目录是否存在
	strUploadPath = Server.MapPath("upload\xls")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	
	fileExt=LCase(file.FileExt)
	If fileExt <> "xls" And fileExt <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errstring = "所选择的不是 Excel 文件！"
	Else
		' 生成日期格式文件名
		fileid = FormatDateTime(Now(),1)&Int(Timer)
		strDestFile = fileid&"."&fileExt
		strDestPath = Server.MapPath("upload")&"\xls\"&strDestFile
		byteFileSize = file.FileSize
		' 保存
		file.SaveAs strDestPath
	End If
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入自EXCEL文件</title>
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入自EXCEL文件</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p><%=byteFileSize%> 字节已上传，正在导入专家信息...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim fieldValue(6)
		Dim sql,conn,connOrigin,result,rsExp,rsTea,rsTmp
		Dim numNewTeacher,numNewExTeacher,numUpdTeacher
		Dim bIsUpdated,bIsInschool
		Dim arrRet(1)
		Dim py,username
		Dim s,i
		numUpdTeacher=0
		numNewTeacher=0
		numNewExTeacher=0
		Set py=New PinyinQuery
		Connect conn
		ConnectOriginDb connOrigin
		sql="SELECT * FROM TEST_THESIS_REVIEW_EXPERT_INFO"
		GetRecordSet conn,rsExp,sql,result
		sql="SELECT * FROM TEACHER_INFO"
		GetRecordSet connOrigin,rsTea,sql,result
		Do While Not rs.EOF
			If IsNull(rs(0).Value) Then Exit Do
			bIsUpdated=False
			' 是否校外专家
			If IsNull(rs(6).Value) Or rs(6).Value="否" Then
				bIsInschool=True
			Else
				bIsInschool=False
			End If
			
			' 姓名
			s=Trim(rs(0).Value)
			If bIsInschool Then i=getTeacherIdByName(s)
			If bIsInschool And i=-1 Then
				bError=True
				errMsg=errMsg&""""&s&"""按校内教师导入，但该教师不存在。"&vbNewLine
			Else
				fieldValue(0)=s
				' 职称
				fieldValue(1)=rs(1).Value
				' 学科专长
				fieldValue(2)=rs(2).Value
				' 单位（住址）
				fieldValue(3)=rs(3).Value
				' 联系方式
				fieldValue(4)=rs(4).Value
				' 邮箱
				fieldValue(5)=rs(5).Value
				' 备注
				fieldValue(6)=rs(6).Value
				
				sql="EXPERT_NAME='"&fieldValue(0)&"' AND INSCHOOL="&Abs(Int(bIsInschool))
				rsExp.Filter=sql
				If rsExp.EOF Then	' 添加记录
					rsExp.AddNew()
					If Not bIsInschool Then	' 添加校外专家到教师信息表
						rsTea.AddNew()
						' 生成登录用户名，并保证不与已有用户名相同
						username=py.getNamePinyinOf(fieldValue(0))
						s=username
						i=0
						Do
							If i>0 Then	s=username&i
							sql="SELECT TEACHERID FROM TEACHER_INFO WHERE TEACHERNO="&toSqlString(s)
							GetRecordSetNoLock connOrigin,rsTmp,sql,result
							CloseRs rsTmp
							i=i+1
						Loop While result>0
						rsTea("TEACHERNO").Value=s
						rsTea("TEACHERNAME").Value=fieldValue(0)
						rsTea("USER_PASSWORD").Value="12345678" ' generatePassword()
						rsTea("IFTEACHER").Value=3
						rsTea("Office_Address").Value=fieldValue(3)
						rsTea("MOBILE").Value=fieldValue(4)
						rsTea("EMAIL").Value=fieldValue(5)
						rsTea("PRO_DUTYID").Value=18	' 职称为其他
						rsTea.Update()
						numNewExTeacher=numNewExTeacher+1
					End If
					numNewTeacher=numNewTeacher+1
				Else							' 更新记录
					If Not bIsInschool Then	' 更新校外专家在教师信息表中的记录
						rsTea.Find("TEACHERID="&rsExp("TEACHER_ID").Value)
						If Not rsTea.EOF Then
							rsTea("USER_PASSWORD").Value="12345678" ' generatePassword()
							rsTea("Office_Address").Value=fieldValue(3)
							rsTea("MOBILE").Value=fieldValue(4)
							rsTea("EMAIL").Value=fieldValue(5)
							rsTea("VALID").Value=0
							rsTea.Update()
						End If
					End If
					numUpdTeacher=numUpdTeacher+1
				End If
				rsExp("EXPERT_NAME").Value=fieldValue(0)
				rsExp("PRO_DUTY_NAME").Value=fieldValue(1)
				rsExp("EXPERTISE").Value=fieldValue(2)
				rsExp("WORKPLACE").Value=fieldValue(3)
				rsExp("MOBILE").Value=fieldValue(4)
				rsExp("EMAIL").Value=fieldValue(5)
				rsExp("MEMO").Value=fieldValue(6)
				rsExp("INSCHOOL").Value=Abs(bIsInschool)
				rsExp.Update()
			End If
			rs.MoveNext()
		Loop
		CloseRs rsTea
		CloseRs rsExp
		' 调用绑定教师ID的存储过程
		conn.Execute "EXEC spUpdateThesisReviewExpertTid"
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
	filepath=Server.MapPath("upload/xls/"&filename)
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&filepath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
	connExcel.Open connstring
	
	Set rs=connExcel.OpenSchema(adSchemaTables)
	Do While Not rs.EOF
		If rs("TABLE_TYPE").Value="TABLE" Then
			table_name=rs("TABLE_NAME").Value
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