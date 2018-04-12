<%Response.Charset="utf-8"%>
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="appgen.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 文件选择页面
	reportNameFmt="\$stu_name_\$stu_no_.+\.(pdf|mht)"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/upload.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>导入论文查重信息自EXCEL文件</b><br>
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择论文查重信息 Excel 文件：<input type="file" name="excelFile" size="100" title="论文查重信息表" /><br />
请选择检测报告 RAR 文件：<input type="file" name="rarFile" size="100" title="检测报告压缩文件" /><br />
检测报告文件名格式（不建议更改）：<input type="text" name="reportNameFmt" size="100" value="<%=reportNameFmt%>" /><br />
<a href="upload/thesisinf_template.xlsx" target="_blank">点击下载论文查重信息表格模板</a><br />
<input type="submit" name="btnupload" value="上 传" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center></body>
<script type="text/javascript">
$(document).ready(function(){
	$('form').submit(function() {
		var valid=checkIfExcel(this.excelFile)&&checkIfRar(this.rarFile);
		if(valid) {
			$(':submit').val("正在上传，请稍候...").attr('disabled',true);
		}
		return valid;
	});
	$(':submit').attr('disabled',false);
});
</script></body></html><%
Case 2	' 上传进程

	Dim fso:Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Dim Upload,tablefile,rarfile,reportNameFmt
	
	Set Upload=New ExtendedRequest
	Set tablefile=Upload.File("excelFile")
	Set rarfile=Upload.File("rarFile")
	reportNameFmt=Upload.Form("reportNameFmt")
	
	' 检查上传目录是否存在
	uploadTablePath = Server.MapPath("upload/xls")
	If Not fso.FolderExists(uploadTablePath) Then fso.CreateFolder(uploadTablePath)
	reportDir = getDateTimeId(Now)
	uploadRarPath = Server.MapPath(reportBaseDir&reportDir)
	If Not fso.FolderExists(uploadRarPath) Then fso.CreateFolder(uploadRarPath)
	
	tableFileExt=LCase(tablefile.FileExt)
	rarFileExt=LCase(rarfile.FileExt)
	If tableFileExt <> "xls" And tableFileExt <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errstring = "论文查重信息表不是 Excel 文件！"
	ElseIf rarFileExt <> "rar" Then
		bError = True
		errstring = "打包的检测报告不是 RAR 压缩文件！"
	ElseIf Len(reportNameFmt)=0 Then
		bError = True
		errstring = "请输入检测报告文件名格式！"
	Else
		' 生成日期格式文件名
		fileid = FormatDateTime(Now(),1)&Int(Timer)
		strDestTableFile = fileid&"."&tableFileExt
		strDestTablePath = "upload/xls/"&strDestTableFile
		byteFileSize = tablefile.FileSize
		' 保存
		tablefile.SaveAs Server.MapPath(strDestTablePath)
		
		strDestRarFile = fileid&"."&rarFileExt
		byteFileSize = byteFileSize+rarfile.FileSize
		rarfile.SaveAs uploadRarPath&"\"&strDestRarFile
	End If
	Set tablefile=Nothing
	Set rarfile=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入论文查重信息自EXCEL文件</title>
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><br /><b>导入论文查重信息自EXCEL文件</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="tableFilepath" value="<%=strDestTablePath%>" />
<input type="hidden" name="rarFilename" value="<%=strDestRarFile%>" />
<input type="hidden" name="reportDir" value="<%=reportDir%>" />
<input type="hidden" name="reportNameFmt" value="<%=reportNameFmt%>" />
<p><%=byteFileSize%> 字节已上传，正在导入论文查重信息和关联检测报告...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errstring%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim sql,sql2,conn,result,rsReview
		Dim thesis_id,stu_id,reproduct_ratio,detect_count,new_status,numThesis
		Dim stu_no,stu_name
		Dim reportFilePath,reportFilename,bFileExists
		Dim file,folder
		Dim regExp:Set regExp=New RegExp
		Dim rag:Set rag=New ReviewAppGen
		
		Randomize
		regExp.IgnoreCase=True
		numThesis=0
		Set folder=fso.GetFolder(Server.MapPath(reportBaseDir&reportDir))
		Connect conn
		Do While Not rs.EOF
			If IsNull(rs(0).Value) Then Exit Do
			stu_name=rs(0).Value
			stu_no=rs(1).Value
			reproduct_ratio=rs(3).Value
			If Right(reproduct_ratio,1)="%" Then	' 复制比为文本格式
				reproduct_ratio=Left(reproduct_ratio,Len(reproduct_ratio)-1)
			ElseIf IsNumeric(reproduct_ratio) Then
				If reproduct_ratio<1 Then	' 复制比为百分比格式
					reproduct_ratio=reproduct_ratio*100
				End If
			End If
			reportFilename=Replace(Replace(reportNameFmt,"\$stu_name",stu_name),"\$stu_no",stu_no)
			regExp.Pattern=reportFilename
			bFileExists=False
			For Each file In folder.Files
				If regExp.Test(file.Name) Then
					bFileExists=True
					Exit For
				End If
			Next
			If Not IsNumeric(reproduct_ratio) Then
				bError=True
				errMsg=errMsg&"学生"""&stu_name&"""的论文复制比为无效值。"&vbNewLine
			End If
			If Not bFileExists Then
				bError=True
				errMsg=errMsg&"找不到学生"""&stu_name&"""的检测报告文件。"&vbNewLine
			End If
			If Not bError Then
				reportFilePath=reportDir&file.Name
				stu_no=toSqlString(rs(1).Value)
				sql="SELECT ID,STU_ID,STU_NAME,SPECIALITY_NAME,THESIS_SUBJECT,THESIS_FILE,REVIEW_APP_EVAL,DETECT_COUNT,TUTOR_ID,TUTOR_NAME FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_NO="&stu_no
				GetRecordSet conn,rsReview,sql,result
				If rsReview.EOF Then
					bError=True
					errMsg=errMsg&"学号不存在:"""&stu_no&"""。"&vbNewLine
				Else
					thesis_id=rsReview("ID").Value
					stu_id=rsReview("STU_ID").Value
					thesis_file=rsReview("THESIS_FILE").Value
					detect_count=rsReview("DETECT_COUNT").Value
					If reproduct_ratio<=10 Then	' 通过
						If detect_count>=1 Then	' 二次检测
							new_status=rsRedetectPassed
						Else
							new_status=rsAgreeReview
							
							' 生成送审申请表
							Dim author:author=rsReview("STU_NAME").Value
							Dim tutor_info:tutor_info=rsReview("TUTOR_NAME").Value&" "&getProDutyNameOf(rsReview("TUTOR_ID").Value)
							Dim speciality:speciality=rsReview("SPECIALITY_NAME").Value
							Dim subject:subject=rsReview("THESIS_SUBJECT").Value
							Dim eval_text:eval_text=rsReview("REVIEW_APP_EVAL").Value
							Dim review_time:review_time=rsReview("SUBMIT_REVIEW_TIME").Value
							If IsNull(review_time) Then review_time=Now
							Dim filename:filename=FormatDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)&".docx"
							Dim filepath:filepath=Server.MapPath("/ThesisReview/tutor/export")&"\"&filename
							rag.Author=author
							rag.StuNo=stu_no
							rag.TutorInfo=tutor_info
							rag.Spec=speciality
							rag.Date=FormatDateTime(review_time,1)
							rag.Subject=subject
							rag.EvalText=eval_text
							rag.ReproductRatio=reproduct_ratio
							If rag.generateApp(filepath)=0 Then
								bError=True
								errMsg=errMsg&"为学生"""&stu_name&"""的论文生成送审申请表时出错。"&vbNewLine
							Else
								sql2=sql2&"UPDATE TEST_THESIS_REVIEW_INFO SET REVIEW_APP="&toSqlString(filename)&" WHERE STU_ID="&stu_id&";"
							End If
						End If
					Else												' 不通过
						If detect_count>=1 Then	' 二次检测
							new_status=rsRedetectUnpassed
						Else
							new_status=rsDetectUnpassed
						End If
					End If
					sql2=sql2&"UPDATE TEST_THESIS_REVIEW_INFO SET REPRODUCTION_RATIO="&reproduct_ratio&",DETECT_REPORT="&toSqlString(reportFilePath)&",REVIEW_STATUS="&new_status&" WHERE STU_ID="&stu_id&";"
					sql2=sql2&"EXEC spAddDetectResult "&thesis_id&","&toSqlString(thesis_file)&","&toSqlString(Now)&","&toSqlString(reportFilePath)&","&reproduct_ratio&";"
					numThesis=numThesis+1
				End If
			End If
			CloseRs rsReview
			rs.MoveNext()
		Loop
		If Len(sql2) Then
			conn.Execute sql2
		End If
		CloseConn conn
		Set rag=Nothing
		Set file=Nothing
		Set flder=Nothing
		Set regExp=Nothing
		addData=numThesis
	End Function
	
	Dim bError,errMsg,ret
	
	tableFilepath=Server.MapPath(Request.Form("tableFilepath"))
	rarFilename=Request.Form("rarFilename")
	reportDir=Request.Form("reportDir")&"/"
	reportNameFmt=Request.Form("reportNameFmt")
	
	' 打包文件
	Dim sourcefile
	Dim rarExe,cmd,wsh
	Dim numFailed,numSucceeded
	
	numFailed=0
	numSucceeded=0
	rarExe=Server.MapPath("rar/Rar.exe")
	sourcefile=Server.MapPath(reportBaseDir&reportDir&rarFilename)
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	Set wsh=Server.CreateObject("WScript.Shell")
	' 解压缩
	cmd=""""&rarExe&""" e -o+ """&sourcefile&""" """&Server.MapPath(reportBaseDir&reportDir)&""""
	Set exec=wsh.Exec(cmd)
	exec.StdOut.ReadAll()
	Set exec=Nothing
	Set wsh=Nothing
	
	' 导入送检结果
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&tableFilepath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
	connExcel.Open connstring
	
	Set rs=connExcel.OpenSchema(adSchemaTables)
	Do While Not rs.EOF
		If rs("TABLE_TYPE")="TABLE" Then
			table_name=rs("TABLE_NAME")
			If InStr("Sheet1$",table_name) Then Exit Do
		End If
		rs.MoveNext()
	Loop
	sql="SELECT * FROM ["&table_name&"A2:D]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	ret=addData()
	CloseRs rs
	CloseConn connExcel
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其他 <%=ret%> 条记录已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=ret%> 条记录已导入。");
<%End If
%>location.href="thesisList.asp";
</script><%
End Select
Set fso=Nothing
%>