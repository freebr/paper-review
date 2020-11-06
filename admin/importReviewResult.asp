<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/automation/ReviewDocumentReader.inc"-->
<!--#include file="../inc/automation/ReviewDocumentWriter.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")

zipUploadPath = Server.MapPath(uploadBasePath(usertypeAdmin, "review_result"))
ensurePathExists zipUploadPath

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 文件选择页面
	activity_id=toUnsignedInt(Request.Form("In_ActivityId2"))
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入送审论文评阅结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery", "upload" %>
</head>
<body>
<center><font size=4><b>导入送审论文评阅结果</b><br />
<form id="fmUpload" action="?step=2" method="POST" enctype="multipart/form-data">
<p>请选择打包评阅书的 RAR 或 ZIP 文件：<input type="file" name="zipFile" size="100" title="评阅书打包文件" /></p>
<p><input type="submit" name="btnsubmit" value="提 交" />&nbsp;
<input type="button" name="btnret" value="返 回" onclick="history.go(-1)" /></p></form></center></body>
<script type="text/javascript">
	$(document).ready(function(){
		$("form").submit(function() {
			var valid=checkIfRarZip(this.zipFile);
			if(valid) {
				$(":submit").val("正在提交，请稍候...").attr("disabled",true);
			}
			return valid;
		});
		$(":submit").attr("disabled",false);
	});
</script></body></html><%
Case 2	' 上传进程

	Dim Upload,zip_file
	
	Set Upload=New ExtendedRequest
	Set zip_file=Upload.File("zipFile")
	activity_id=toUnsignedInt(Upload.Form("In_ActivityId"))
	
	zipFileExt=LCase(zip_file.FileExt)
	If activity_id="0" Then
		bError = True
		errMsg = "请选择评阅活动！"
	ElseIf zipFileExt <> "rar" And zipFileExt <> "zip" Then
		bError = True
		errMsg = "上传文件必须为 RAR 或 ZIP 压缩文件！"
	Else
		destFile = timestamp()&"."&zipFileExt
		zip_file.SaveAs resolvePath(zipUploadPath,destFile)
	End If
	Set zip_file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>导入送审论文评阅结果</title>
<% useStylesheet "admin" %>
<% useScript "jquery" %>
</head>
<body>
<center><br /><b>导入送审论文评阅结果</b><br /><br /><%
	If Not bError Then %>
<form id="fmUploadFinish" action="?step=3" method="POST">
<input type="hidden" name="filename" value="<%=destFile%>" />
<p>文件上传成功，正在导入送审论文评阅结果...</p>
<p align="left"><span id="output" class="output-message"></span></p></form>
<script type="text/javascript">
	var progfile=location.origin+"<%=resolvePath(tempPath(),"prog_"&Session("id")&".txt")%>";
	setTimeout(function() {
		$("#fmUploadFinish").submit();
		window.pending=true;
		window.pendingOperation="导入送审论文评阅结果";
		setTimeout(refreshProgress, 500);
	});
	function refreshProgress() {
		$.get(progfile, function (data, status) {
			if(status=="success") {
				$("#output").html(data);
				if(data.match(/\t$/)) {
					$(":submit").val("提 交").attr("disabled", false);
				} else {
					setTimeout(refreshProgress, 500);
				}
			} else {
				setTimeout(refreshProgress, 500);
			}
		});
	}
</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
Case 3	' 数据读取，导入到数据库

	Function addData(upload_path, authors)
		' 添加数据
		outputMessage "读取评阅书信息……"
		Dim folder: Set folder = fso.GetFolder(upload_path)
		Dim file
		Dim reader: Set reader = New ReviewDocumentReader
		Dim sql,conn,rsa,rsb
		Dim review_count:review_count=0
		Dim new_msg
		Dim import_time:import_time=Now
		Dim wd:Set wd = Server.CreateObject("Word.Application")
		ConnectDb conn
		Set authors=CreateDictionary()
		For Each file In folder.Files
			If fso.FileExists(file.Path) And _
				isMatched(fso.getExtensionName(file.Name), "^docx?$",True) Then
				Dim paper_id,author,tutorinfo,subject,speciality
				Dim researchway,review_type,paper_type_name,stu_type,stu_type_name
				Dim reviewer1,reviewer2,expert_id:expert_id=-1
				
				outputMessage Format("正在处理：评阅书[{0}]……",file.Name)
				reader.extractInfoFromReviewDocument resolvePath(upload_path,file.Name),wd
				' 学号
				outputMessage Format("正在导入：学号[{0}]的评阅记录……",reader.StuNo)
				
				sql=Format("SELECT ID,STU_NAME,THESIS_SUBJECT,TUTOR_ID,TUTOR_NAME,SPECIALITY_NAME,RESEARCHWAY_NAME,REVIEW_TYPE,THESIS_FORM,TEACHTYPE_ID,TEACHTYPE_NAME,REVIEWER1,REVIEWER2 FROM ViewDissertations WHERE STU_NO={0}",toSqlString(reader.StuNo))
				Set rsa=conn.Execute(sql)
				If rsa.EOF Then
					bError=True
					new_msg=Format("学号不存在：[{0}]。",reader.StuNo)
					errMsg=errMsg&new_msg&vbNewLine
					outputMessage new_msg
				Else
					' 论文基本信息
					paper_id=rsa("ID")
					author=rsa("STU_NAME")
					tutorinfo=rsa("TUTOR_NAME") '&" "&getProDutyNameOf(rsa("TUTOR_ID"))
					subject=rsa("THESIS_SUBJECT")
					speciality=rsa("SPECIALITY_NAME")
					researchway=rsa("RESEARCHWAY_NAME")
					review_type=rsa("REVIEW_TYPE")
					paper_type_name=rsa("THESIS_FORM")
					stu_type=rsa("TEACHTYPE_ID")
					stu_type_name=rsa("TEACHTYPE_NAME")
					reviewer1=rsa("REVIEWER1")
					reviewer2=rsa("REVIEWER2")
					CloseRs rsa
					' 专家ID
					If Len(reader.ExpertTel) Then
						expert_id=getExpertIdByTelephone(reader.ExpertTel)
					End If
					If expert_id=-1 Then
						expert_id=getExpertIdByName(reader.ExpertName)
					End If
					If expert_id=-1 Then
						If Right(fso.GetBaseName(file.Name),2) = "_2" Then
							expert_id=reviewer2
						Else
							expert_id=reviewer1
						End If
					End If
					If expert_id=-1 Then
						bError=True
						new_msg=Format("为学生[{0}]给出以下评阅书的专家未录入系统评阅专家库：{1}。",author,file.Name)
						errMsg=errMsg&new_msg&vbNewLine
						outputMessage new_msg
					Else
						Dim expert_name,expert_pro_duty,expert_expertise,expert_workplace
						Dim expert_address,expert_mailcode,expert_telephone,expert_mobile
						Dim display_status:display_status=0	' 评阅书状态默认为不开放
						
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

						outputMessage Format("学生[{0}]<-->专家[{1}]，送审日期：{2}", author,expert_name,toDateTime(reader.ReviewTime,1))
						outputMessage "正在生成评阅书……"
						
						Dim rg:Set rg=New ReviewDocumentWriter
						Dim scores
						rg.Author=author
						rg.StuNo=reader.StuNo
						rg.TutorInfo=tutorinfo
						rg.Subject=subject
						rg.ResearchWay=researchway
						rg.Date=toDateTime(reader.ReviewTime,1)
						rg.ReviewPattern=reader.ReviewPattern
						rg.ExpertName=expert_name
						rg.ExpertProDuty=isZeroString(reader.ExpertProDuty,expert_pro_duty)
						rg.ExpertExpertise=isZeroString(reader.ExpertExpertise,expert_expertise)
						rg.ExpertWorkplace=isZeroString(reader.ExpertWorkplace,expert_workplace)
						rg.ExpertAddress=isZeroString(reader.ExpertAddress,expert_address)
						rg.ExpertTel1=isZeroString(reader.ExpertTel,expert_telephone)
						rg.ExpertTel2=expert_mobile
						rg.ExpertMasterLevel=reader.ExpertMasterLevel
						rg.Comment=reader.Comment
						rg.Suggestion=reader.Suggestion
						rg.CorrelationLevel=reader.CorrelationLevel
						rg.ReviewResult=reader.ReviewResult
						rg.ReviewLevel=reader.ReviewLevel
						rg.PaperTypeName=paper_type_name
						If stu_type=5 Or stu_type=6 Then	' MEM/MBA评阅书，计算评价指标总分
							rg.Spec=speciality
							scores=Join(reader.Scores,",")
							rg.Scores=scores
							If IsNull(reader.ScoreParts) Then
								rg.ScoreParts=Null
							Else
								rg.ScoreParts=Join(reader.ScoreParts,",")
							End If
							rg.TotalScore=reader.TotalScore
						Else
							scores=Null
							rg.Scores=Null
							rg.ScoreParts=Null
						End If
						Dim template_file,filename,review_file_paths

						sql="SELECT REVIEW_FILE FROM ReviewTypes WHERE ID="&review_type
						Set rsb = conn.Execute(sql)
						If rsb.EOF Then
							bError=True
							new_msg="评阅书模板丢失，无法完成评阅操作，请联系系统管理员。"
							errMsg=errMsg&new_msg&vbNewLine
							outputMessage new_msg
							Set rg=Nothing
							CloseRs rsb
							Exit For
						End If
						template_file=Server.MapPath(resolvePath(uploadBasePath(usertypeAdmin,"review_template"),rsb(0)))
						CloseRs rsb

						filename=toDateTime(import_time,1)&Int(Timer)&Int(Rnd()*999)
						review_file_paths=reviewFileVersionPath(filename)
						' 生成评阅书
						bError=rg.exportReviewDocument(review_file_paths(0),review_file_paths(1),review_file_paths(2),template_file,stu_type,wd)=0
						Set rg=Nothing
						outputMessage Format("评阅书已导出，编号：{0}",filename)
						
						' 插入评阅记录
						sql="EXEC spAddReviewRecord ?,?,?,?,?,?,?,?,?,?,?,?,?,?"
						Dim ret:Set ret=ExecQuery(conn,sql,_
							CmdParam("paper_id",adInteger,4,paper_id),_
							CmdParam("reviewer_id",adInteger,4,expert_id),_
							CmdParam("reviewer_master_level",adInteger,4,reader.ExpertMasterLevel),_
							CmdParam("score_data",adVarWChar,500,scores),_
							CmdParam("comment",adLongVarWChar,5000,isZeroString(reader.Comment,"")),_
							CmdParam("suggestion",adLongVarWChar,5000,isZeroString(reader.Suggestion,"")),_
							CmdParam("correlation_level",adInteger,4,reader.CorrelationLevel),_
							CmdParam("overall_rating",adInteger,4,reader.ReviewLevel),_
							CmdParam("defence_opinion",adInteger,4,reader.ReviewResult),_
							CmdParam("review_time",adDate,4,reader.ReviewTime),_
							CmdParam("review_pattern",adVarWChar,100,reader.ReviewPattern),_
							CmdParam("review_file",adVarWChar,50,filename),_
							CmdParam("display_status",adInteger,4,display_status),_
							CmdParam("creator",adInteger,4,Session("Id")))
						If Not authors.Exists(author) Then authors(author)=0
						authors(author)=authors(author)+1
						review_count=review_count+1
					End If
				End If
				On Error Resume Next
				fso.DeleteFile file.Path
				On Error GoTo 0
			End If
		Next
		fso.DeleteFolder upload_path
		CloseConn conn
		wd.Quit()
		Set wd=Nothing
		Set file = Nothing
		Set folder = Nothing
		addData=review_count
	End Function

	Function outputMessage(msg)
		On Error Resume Next
		streamLog.WriteText Format("[{0}]{1}", toDateTime(Now, 3), msg) & vbNewLine
		streamLog.SaveToFile progfile, 2
		streamLog.Position=streamLog.Size
		On Error GoTo 0
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

	Function getReviewLevelId(name, stu_type_name)
		getReviewLevelId = dictReviewLevelId(stu_type_name)(name)
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
	
	Function getReviewResultId(name, stu_type_name)
		getReviewResultId = dictReviewResultId(stu_type_name)(name)
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

	Server.ScriptTimeout=3600
	Dim fso:Set fso=CreateFSO()
	Dim bError,errMsg
	Dim progfile,streamLog
	Dim authors
	progfile=Server.MapPath(resolvePath(tempPath(),"prog_"&Session("Id")&".txt"))
	If fso.FileExists(progfile) Then fso.DeleteFile progfile

	Set streamLog=Server.CreateObject("ADODB.Stream")
	streamLog.Mode=3
	streamLog.Type=2
	streamLog.Open()

	' 解压缩
	outputMessage "解压缩打包文件……"
	zipFilename=Request.Form("filename")
	extractPath=resolvePath(zipUploadPath,fso.GetBaseName(zipFilename))
	If fso.FolderExists(extractPath) Then fso.DeleteFolder extractPath
	fso.CreateFolder extractPath
	extractFile resolvePath(zipUploadPath,zipFilename), extractPath

	' 添加数据
	ret=addData(extractPath, authors)
	
	outputMessage "导入完成。"&Chr(9)
	streamLog.Close()
	Set streamLog=Nothing
	If fso.FileExists(progfile) Then fso.DeleteFile progfile
	If authors.Count Then
		Dim authors_list
		For Each key In authors
			If Len(authors_list) Then authors_list = authors_list & ","
			authors_list = authors_list & Format("{0} {1} 份", key, authors(key))
		Next
		logtxt = Format("教务员[{0}]为以下学生导入送审论文评阅结果（共计 {1} 份评阅书）：{2}。", Session("name"), ret, authors_list)
		writeLog logtxt
	End If
	Server.ScriptTimeout=90
%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其他 <%=ret%> 个评阅结果已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%	Else %>
	alert("操作成功，<%=ret%> 个评阅结果已导入。");
<%	End If
%>location.href="paperList.asp";
</script><%
End Select
Set fso = Nothing
%>