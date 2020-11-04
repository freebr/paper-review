<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/automation/ReviewApplicationFormWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

reportDir = getDateTimeId(Now)
tableUploadPath = Server.MapPath(uploadBasePath(usertypeAdmin, "detect_stats"))
zipUploadPath = Server.MapPath(resolvePath(uploadBasePath(usertypeAdmin, "detect_report"),reportDir))
ensurePathExists tableUploadPath
ensurePathExists zipUploadPath

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
	report_name_format="@stu_name_@stu_no_.+\.(pdf|mht|htm(l?))"
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入送检论文查重信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body>
<center><font size=4><b>导入送检论文查重信息</b><br>
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<input type="file" name="tableFile" size="100" title="送检论文查重信息表" /></p>
<p>请选择检测报告 RAR 或 ZIP 文件：<input type="file" name="zipFile" size="100" title="检测报告打包文件" /></p>
<p>检测报告文件名格式（不建议更改）：<input type="text" name="reportNameFmt" size="100" value="<%=report_name_format%>" /></p>
<p><a href="upload/paperinf_template.xlsx" target="_blank">点击下载送检论文查重信息表格模板</a></p>
<p><input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center></body>
<script type="text/javascript">
	$(document).ready(function(){
		$('form').submit(function() {
			var valid=checkIfExcel(this.tableFile)&&checkIfRarZip(this.zipFile);
			if(valid) {
				$(':submit').val("正在提交，请稍候...").attr('disabled',true);
			}
			return valid;
		});
		$(':submit').attr('disabled',false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim Upload,table_file,zip_file,report_name_format
	
	Set Upload=New ExtendedRequest
	Set table_file=Upload.File("tableFile")
	Set zip_file=Upload.File("zipFile")
	report_name_format=Upload.Form("reportNameFmt")
	
	tableFileExt=LCase(table_file.FileExt)
	zipFileExt=LCase(zip_file.FileExt)
	If tableFileExt <> "xls" And tableFileExt <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errMsg = "送检论文查重信息表必须为 Excel 文件！"
	ElseIf zipFileExt <> "rar" And zipFileExt <> "zip" Then
		bError = True
		errMsg = "检测报告必须为 RAR 或 ZIP 压缩文件！"
	ElseIf Len(report_name_format)=0 Then
		bError = True
		errMsg = "请输入检测报告文件名格式！"
	Else
		fileid = timestamp()
		destTableFile = fileid&"."&tableFileExt
		table_file.SaveAs resolvePath(tableUploadPath,destTableFile)
		destZipFile = fileid&"."&zipFileExt
		zip_file.SaveAs resolvePath(zipUploadPath,destZipFile)
	End If
	Set table_file=Nothing
	Set zip_file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入送检论文查重信息</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入送检论文查重信息</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="tableFilename" value="<%=destTableFile%>" />
<input type="hidden" name="zipFilename" value="<%=destZipFile%>" />
<input type="hidden" name="reportNameFmt" value="<%=report_name_format%>" />
<p>文件上传成功，正在导入论文查重信息和关联检测报告...</p></form>
<script type="text/javascript">
	setTimeout(function() {
		$("#fmUploadFinish").submit();
		window.pending=true;
		window.pendingOperation="导入论文查重信息和关联检测报告";
	});
</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim conn,count,rsReview
		Dim sql,sql2
		Dim sql3:sql3="DECLARE @comment nvarchar(MAX)"
		Dim countDetect,newStatus,reproductRatio,reviewStatus,stuId,stuName,stuNo,paperFile,paperId
		Dim numPapers
		Dim willMakeApp:willMakeApp=False
		Dim reportFilePath,reportFilename,reportFound
		Dim fso,file,folder: Set fso=CreateFSO()
		Dim regExp:Set regExp=New RegExp
		Dim rag:Set rag=New ReviewApplicationFormWriter
		Dim importTime:importTime=Now
		
		Randomize()
		regExp.IgnoreCase=True
		numPapers=0
		Set folder=fso.GetFolder(zipUploadPath)
		Connect conn
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			stuName=rs(0)
			stuNo=rs(1)
			reproductRatio=rs(3)
			If Right(reproductRatio,1)="%" Then	' 复制比为文本格式
				reproductRatio=Left(reproductRatio,Len(reproductRatio)-1)
			ElseIf IsNumeric(reproductRatio) Then
				If reproductRatio<1 Then	' 复制比为百分比格式
					reproductRatio=reproductRatio*100
				End If
			End If
			If Not IsNumeric(reproductRatio) Then
				bError=True
				errMsg=errMsg&"学生["&stuName&"]的论文复制比为无效值。"&vbNewLine
			Else
				reportFilename=Replace(Replace(report_name_format,"@stu_name",stuName),"@stu_no",stuNo)
				regExp.Pattern=reportFilename
				reportFound=False
				For Each file In folder.Files
					If regExp.Test(file.Name) Then
						reportFound=True
						Exit For
					End If
				Next
				If Not reportFound Then
					bError=True
					errMsg=errMsg&"找不到学生["&stuName&"]的检测报告文件。"&vbNewLine
				Else
					reportFilePath=resolvePath(reportDir,file.Name)
					sql="SELECT ID,STU_ID,STU_NAME,SPECIALITY_NAME,THESIS_SUBJECT,THESIS_FILE,REVIEW_STATUS,REVIEW_APP_EVAL,SUBMIT_REVIEW_TIME,DETECT_COUNT,TUTOR_ID,TUTOR_NAME FROM ViewDissertations WHERE STU_NO="&toSqlString(stuNo)
					GetRecordSet conn,rsReview,sql,count
					If rsReview.EOF Then
						bError=True
						errMsg=errMsg&"学号不存在:"""&stuNo&"""。"&vbNewLine
					Else
						paperId=rsReview("ID")
						stuId=rsReview("STU_ID")
						stuName=rsReview("STU_NAME")
						reviewStatus=rsReview("REVIEW_STATUS")
						paperFile=rsReview("THESIS_FILE")
						countDetect=rsReview("DETECT_COUNT")
						If reviewStatus=rsAgreedDetect Then
							If reproductRatio<=10 Then	' 通过
								newStatus=rsAgreedReview
								willMakeApp=True
							Else	' 不通过
								If countDetect=0 Then
									newStatus=rsDetectUnpassed
								Else	' 二次检测
									newStatus=rsRedetectUnpassed
								End If
							End If
						Else	' 更新为与当前检测次数一致的状态
							If reproductRatio<=10 Then	' 通过
								newStatus=rsAgreedReview
								willMakeApp=True
							Else	' 不通过
								If countDetect=1 Then
									newStatus=rsDetectUnpassed
								Else	' 二次检测
									newStatus=rsRedetectUnpassed
								End If
							End If
						End If
						If willMakeApp Then
							' 生成送审申请表
							Dim author:author=stuName
							Dim tutor_id:tutor_id=rsReview("TUTOR_ID")
							Dim tutor_info:tutor_info=rsReview("TUTOR_NAME")&" "&getProDutyNameOf(tutor_id)
							Dim speciality:speciality=rsReview("SPECIALITY_NAME")
							Dim subject:subject=rsReview("THESIS_SUBJECT")
							Dim comment:comment=rsReview("REVIEW_APP_EVAL")
							Dim auditTime:auditTime=rsReview("SUBMIT_REVIEW_TIME")
							If IsNull(auditTime) Then auditTime=importTime
							Dim filename:filename=FormatDateTime(auditTime,1)&Int(Timer)&Int(Rnd()*999)&".docx"
							Dim filepath:filepath=Server.MapPath(exportBasePath(usertypeTutor)&filename)
							rag.Author=author
							rag.StuNo=stuNo
							rag.TutorInfo=tutor_info
							rag.Spec=speciality
							rag.Date=FormatDateTime(auditTime,1)
							rag.Subject=subject
							rag.Comment=comment
							rag.ReproductRatio=reproductRatio
							If rag.generateApp(filepath)=0 Then
								bError=True
								errMsg=errMsg&"为学生["&stuName&"]的论文生成送审申请表时出错。"&vbNewLine
							Else
								sql2=sql2&"UPDATE Dissertations SET REVIEW_APP="&toSqlString(filename)&" WHERE STU_ID="&stuId&";"
								sql3=sql3&Format("SET @comment={0};EXEC spAddAuditRecord {1},{2},{3},{4},{5},1,@comment;",_
									toSqlString(comment),paperId,toSqlString(filename),auditTypeReviewApp,toSqlString(auditTime),tutor_id)
							End If
						End If
						sql2=sql2&"UPDATE Dissertations SET REPRODUCTION_RATIO="&reproductRatio&",DETECT_REPORT="&toSqlString(reportFilePath)&",REVIEW_STATUS="&newStatus&" WHERE STU_ID="&stuId&";"
						sql2=sql2&"EXEC spAddDetectResult "&paperId&","&toSqlString(paperFile)&","&toSqlString(importTime)&","&toSqlString(reportFilePath)&","&reproductRatio&",1;"
						numPapers=numPapers+1
					End If
					CloseRs rsReview
				End If
			End If
			rs.MoveNext()
		Loop
		If Len(sql2) Then
			conn.Execute sql2
		End If
		If Len(sql3) Then
			conn.Execute sql3
		End If
		CloseConn conn
		Set rag=Nothing
		Set file=Nothing
		Set folder=Nothing
		Set fso=Nothing
		Set regExp=Nothing
		addData=numPapers
	End Function
	
	Server.ScriptTimeout=600
	Dim bError,errMsg,ret
	
	tableFilePath=resolvePath(tableUploadPath,Request.Form("tableFilename"))
	zipFilename=Request.Form("zipFilename")
	report_name_format=Request.Form("reportNameFmt")
	
	' 打包文件
	numFailed=0
	numSucceeded=0
	' 解压缩
	ExtractFile resolvePath(zipUploadPath,zipFilename), zipUploadPath
	
	' 导入送检结果
	Set connExcel=Server.CreateObject("ADODB.Connection")
	connstring="Provider=Microsoft.ACE.OLEDB.12.0;Data Source="&tableFilePath&";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"""
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
	Server.ScriptTimeout=90
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其他 <%=ret%> 篇论文的检测结果已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%	Else %>
	alert("操作成功，<%=ret%> 篇论文的检测结果已导入。");
<%	End If
%>location.href="paperList.asp";
</script><%
End Select
%>