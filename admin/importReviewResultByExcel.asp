<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/automation/ReviewDocumentWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

tableUploadPath = Server.MapPath(uploadBasePath(usertypeAdmin, "review_result_table"))
ensurePathExists tableUploadPath

tmpDir=Server.MapPath("tmp")
progfile=resolvePath(tmpDir,"prog_"&Session("Id")&".txt")
step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
	activity_id=toUnsignedInt(Request.Form("In_ActivityId2"))
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入评阅结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body>
<center><font size=4><b>导入评阅结果</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择要导入的 Excel 文件：<input type="file" name="tableFile" size="100" /></p>
<p><a href="upload/review_result_template.xlsx" target="_blank">点击下载评阅结果表格模板</a></p>
<p><input type="submit" name="btnsubmit" value="提 交" />&nbsp;
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
	activity_id=toUnsignedInt(Upload.Form("In_ActivityId"))
	task_progress=Upload.Form("In_TASK_PROGRESS")
	review_status=Upload.Form("In_REVIEW_STATUS")
	select_mode=Upload.Form("selectmode")

	' 删除已有临时目录
	If fso.FolderExists(tmpDir) Then fso.DeleteFolder tmpDir
	fso.CreateFolder tmpDir
	' 删除已有临时文件
	If fso.FileExists(progfile) Then fso.DeleteFile progfile

	file_ext=LCase(file.FileExt)
	If activity_id="0" Then
		bError = True
		errMsg = "请选择评阅活动！"
	ElseIf file_ext <> "xls" And file_ext <> "xlsx" Then	' 不被允许的文件类型
		bError = True
		errMsg = "所选择的不是 Excel 文件！"
	Else
		destFile = timestamp()&"."&file_ext
		file.SaveAs resolvePath(tableUploadPath,destFile)
	End If
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入评阅结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入评阅结果</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="filename" value="<%=destFile%>" />
<p>文件上传成功，正在导入评阅结果...</p></form>
<script type="text/javascript">setTimeout("$('#fmUploadFinish').submit()",500);</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData()
		' 添加数据
		Dim sql,conn,rsa,rsb
		Dim paper_ids
		Dim field_count:field_count=rs.Fields.Count
		Dim paper_count:paper_count=0
		Dim wd:Set wd = Server.CreateObject("Word.Application")
		ConnectDb conn
		Do While Not rs.EOF
			If IsNull(rs(0)) Then Exit Do
			Dim paper_id,stu_no,author,tutorinfo,subject,speciality
			Dim researchway,review_type,expert_id
			' 学号
			stu_no=toSqlString(Trim(rs(0)))
			outputMessage Format("正在导入：学号[{0}]的评阅记录……<br/>",stu_no)
			
			sql="SELECT ID,STU_NAME,THESIS_SUBJECT,TUTOR_ID,TUTOR_NAME,SPECIALITY_NAME,RESEARCHWAY_NAME,REVIEW_TYPE,TEACHTYPE_NAME FROM ViewDissertations WHERE STU_NO="&stu_no
			Set rsa=conn.Execute(sql)
			If rsa.EOF Then
				bError=True
				errMsg=errMsg&"学号不存在：["&rs(0)&"]。"&vbNewLine
			Else
				' 论文基本信息
				paper_id=rsa("ID")
				author=rsa("STU_NAME")
				tutorinfo=rsa("TUTOR_NAME")&" "&getProDutyNameOf(rsa("TUTOR_ID"))
				subject=rsa("THESIS_SUBJECT")
				speciality=rsa("SPECIALITY_NAME")
				researchway=rsa("RESEARCHWAY_NAME")
				review_type=rsa("REVIEW_TYPE")
				stu_type=rsa("TEACHTYPE_NAME")
				CloseRs rsa
				' 专家ID
				expert_id=getTeacherIdByName(rs(12))
				If expert_id=-1 Then
					bError=True
					errMsg=errMsg&"学生["&author&"]所匹配的专家["&rs(12)&"]未在评阅专家库。"&vbNewLine
				Else
					Dim expert_name,expert_pro_duty,expert_expertise,expert_workplace
					Dim expert_address,expert_mailcode,expert_telephone,expert_mobile
					Dim scores, overall_rating
					Dim correlation_level,master_level, review_result, review_level
					Dim comment,suggestion,review_time,display_status
					
					sql="SELECT * FROM Experts WHERE TEACHER_ID="&expert_id
					Set rsb=conn.Execute(sql)
					expert_name=rsb("EXPERT_NAME")
					expert_pro_duty=rsb("PRO_DUTY_NAME")
					expert_expertise=rsb("EXPERTISE")
					expert_workplace=rsb("WORKPLACE")
					expert_address=rsb("ADDRESS")
					expert_mailcode=rsb("MAILCODE")
					expert_telephone=rsb("TELEPHONE")
					expert_mobile=rsb("MOBILE")
					CloseRs rsb

					Dim i
					If stu_type=5 Then
						ReDim scores(11)
					ElseIf stu_type=6 Then
						ReDim scores(12)
					End If
					For i=1 To UBound(scores)
						scores(i)=rs(30+i)
					Next

					' 论文评分
					overall_rating=rs(field_count-6)
					' 评阅结果
					review_result=getReviewResultId(rs(14),stu_type)
					' 评阅时间
					review_time=rs(30)
					' 对论文涉及内容的熟悉程度
					master_level=getMasterLevelId(rs(field_count-5))
					' 学位论文内容与申请学位专业的相关性
					correlation_level=getCorrelationLevelId(rs(field_count-4))
					' 对学位论文的总体评价
					review_level=getReviewLevelId(rs(field_count-3),stu_type)
					' 评阅专家对论文的学术评语
					comment=rs(field_count-2)
					' 论文存在的不足之处和建议等
					comment=rs(field_count-1)
					' 评阅书状态
					display_status=0

					outputMessage "正在生成评阅书……"
					
					Dim rg:Set rg=New ReviewDocumentWriter
					rg.Author=author
					rg.TutorInfo=tutorinfo
					rg.Subject=subject
					rg.ResearchWay=researchway
					rg.Date=toDateTime(review_time,1)
					rg.ExpertName=expert_name
					rg.ExpertProDuty=expert_pro_duty
					rg.ExpertExpertise=expert_expertise
					rg.ExpertWorkplace=expert_workplace
					rg.ExpertAddress=expert_address
					rg.ExpertMailcode=expert_mailcode
					rg.ExpertTel1=expert_telephone
					rg.ExpertTel2=expert_mobile
					rg.ExpertMasterLevel=master_level
					rg.Comment=comment
					rg.Suggestion=suggestion
					rg.CorrelationLevel=correlation_level
					rg.ReviewResult=review_result
					rg.ReviewLevel=review_level
					rg.ThesisType=review_type
					If reviewfile_type=2 Then	' ME/MBA评阅书，计算评价指标总分
						rg.Spec=speciality
						rg.Scores=Join(scores,",")
						Dim arrScorePartPower,arrScorePower
						Dim scoreParts
						Dim tmp,code_power1,code_power2
						loadReviewScoringInfo review_type,tmp,code_power1,code_power2
						code_power1=Replace(code_power1,"[","Array(")
						code_power1=Replace(code_power1,"]",")")
						code_power2=Replace(code_power2,"[","Array(")
						code_power2=Replace(code_power2,"]",")")
						arrScorePartPower=Eval(code_power1)
						arrScorePower=Eval(code_power2)
						Dim j,k
						k=0
						For i=0 To UBound(arrScorePartPower)
							Dim partScore:partScore=0
							For j=0 To UBound(arrScorePower(i))
								scores(k)=Int(scores(k))
								partScore=partScore+scores(k)*arrScorePower(i)(j)
								k=k+1
							Next
							If i>0 Then scoreParts=scoreParts&","
							partScore=partScore*arrScorePartPower(i)
							scoreParts=scoreParts&partScore
						Next
						rg.ScoreParts=scoreParts
						rg.TotalScore=overall_rating
					End If
					Dim template_file,filename,review_file_paths,reviewfile_type
					If stu_type=5 Or stu_type=6 Then
						reviewfile_type=2
					Else
						reviewfile_type=1
					End If

					sql="SELECT REVIEW_FILE FROM ReviewTypes WHERE ID="&review_type
					Set rsb = conn.Execute(sql)
					If rsb.EOF Then
						bError=True
						errMsg=errMsg&"评阅书模板丢失，无法完成评阅操作，请联系系统管理员。"
						Set rg=Nothing
						CloseRs rsb
						Exit Do
					End If
					' 生成评阅书
					template_file=Server.MapPath(uploadBasePath(usertypeAdmin,"review_template")&rsb(0))
					CloseRs rsb

					filename=toDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)
					review_file_paths=reviewFileVersionPath(filename)
					bError=rg.exportReviewDocument(review_file_paths(0),review_file_paths(1),review_file_paths(2),template_file,reviewfile_type,wd)=0
					Set rg=Nothing
					outputMessage "完成！<br/>"

					' 插入评阅记录
					review_pattern="凡科送审平台"
					sql="EXEC spAddReviewRecord ?,?,?,?,?,?,?,?,?,?,?,?,?,?"
					Dim ret:Set ret=ExecQuery(conn,sql,_
						CmdParam("paper_id",adInteger,4,paper_id),_
						CmdParam("reviewer_id",adInteger,4,expert_id),_
						CmdParam("reviewer_master_level",adInteger,4,master_level),_
						CmdParam("score_data",adVarWChar,500,scores),_
						CmdParam("comment",adLongVarWChar,5000,comment),_
						CmdParam("suggestion",adLongVarWChar,5000,suggestion),_
						CmdParam("correlation_level",adInteger,4,1),_
						CmdParam("overall_rating",adInteger,4,review_level),_
						CmdParam("defence_opinion",adInteger,4,review_result),_
						CmdParam("review_time",adDate,4,review_time),_
						CmdParam("review_pattern",adVarWChar,100,review_pattern),_
						CmdParam("review_file",adVarWChar,50,filename),_
						CmdParam("display_status",adInteger,4,display_status),
						CmdParam("creator",adInteger,4,Session("Id")))
					Set rsb=ret("rs")
					If rsb(0) >= 2 Then
						' 更新论文状态
						paper_ids=paper_ids&","&paper_id
					End If
					CloseRs rsb
					paper_count=paper_count+1
				End If
			End If
			rs.MoveNext()
		Loop
		If Len(paper_ids) Then
			sql=Format("UPDATE Dissertations SET REVIEW_STATUS={0} WHERE ID IN (0{1})",rsReviewed,paper_ids)
			conn.Execute sql
		End If
		CloseConn conn
		wd.Quit()
		Set wd=Nothing
		addData=paper_count
	End Function

	Function outputMessage(msg)
		streamLog.WriteText msg
		streamLog.SaveToFile progfile,2
		streamLog.Position=streamLog.Size
	End Function

	Function getMasterLevelId(name)
		getMasterLevelId = dictMasterLevelId(name)
	End Function
	Dim dictMasterLevelId:Set dictMasterLevelId=CreateDictionary()
	dictMasterLevelId.Add "优", 1
	dictMasterLevelId.Add "良", 2
	dictMasterLevelId.Add "中", 3

	Function getCorrelationLevelId(name)
		getCorrelationLevelId = dictCorrelationLevelId(name)
	End Function
	Dim dictCorrelationLevelId:Set dictCorrelationLevelId=CreateDictionary()
	dictCorrelationLevelId.Add "相关", 1
	dictCorrelationLevelId.Add "不相关", 2

	Function getReviewLevelId(name, stu_type)
		getReviewLevelId = dictReviewLevelId(stu_type)(name)
	End Function
	Dim dictReviewLevelId:Set dictReviewLevelId=CreateDictionary()
	Dim dictSub:Set dictSub=CreateDictionary()
	dictReviewLevelId.Add "工商管理硕士", dictSub
	dictReviewLevelId.Add "EMBA", dictSub
	dictSub.Add "优", 1
	dictSub.Add "良", 2
	dictSub.Add "中", 3
	dictSub.Add "差", 4
	Set dictSub=CreateDictionary()
	dictReviewLevelId.Add "工程硕士", dictSub
	dictSub.Add "优秀", 1
	dictSub.Add "良好", 2
	dictSub.Add "一般", 3
	dictSub.Add "较差", 4
	Set dictSub=CreateDictionary()
	dictReviewLevelId.Add "会计硕士", dictSub
	dictSub.Add "优秀", 1
	dictSub.Add "良好", 2
	dictSub.Add "一般", 3
	dictSub.Add "较差", 4
	
	Function getReviewResultId(name, stu_type)
		getReviewResultId = dictReviewResultId(stu_type)(name)
	End Function
	Dim dictReviewResultId:Set dictReviewResultId=CreateDictionary()
	Set dictSub=CreateDictionary()
	dictReviewResultId.Add "工商管理硕士", dictSub
	dictReviewResultId.Add "EMBA", dictSub
	dictSub.Add "同意答辩", 1
	dictSub.Add "适当修改后答辩", 2
	dictSub.Add "需做重大修改后方可答辩", 3
	dictSub.Add "未达到答辩要求", 4
	Set dictSub=CreateDictionary()
	dictReviewResultId.Add "工程硕士", dictSub
	dictSub.Add "同意答辩", 1
	dictSub.Add "需对学位论文进行适当修改", 2
	dictSub.Add "需对学位论文进行重大修改", 3
	dictSub.Add "不同意答辩", 4
	Set dictSub=CreateDictionary()
	dictReviewResultId.Add "会计硕士", dictSub
	dictSub.Add "同意答辩", 1
	dictSub.Add "需对学位论文进行适当修改", 2
	dictSub.Add "需对学位论文进行重大修改", 3
	dictSub.Add "不同意答辩", 4

	Dim bError,errMsg
	Dim streamLog

	filename=Request.Form("filename")
	filepath=resolvePath(tableUploadPath,filename)

	Set streamLog=Server.CreateObject("ADODB.Stream")
	streamLog.Mode=3
	streamLog.Type=2
	streamLog.Open()

	outputMessage "正在读取表格数据……<br/>"

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
	sql="SELECT * FROM ["&table_name&"A2:BE]"
	Set rs=connExcel.Execute(sql)
	' 添加数据
	ret=addData()
	CloseRs rs
	CloseConn connExcel

	outputMessage "导入完成。"
	streamLog.Close()
	Set streamLog=Nothing
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余<%=ret%>条记录已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，<%=ret%>条记录已导入。");
<%End If
%>location.href="paperList.asp";
</script><%
End Select
%>