<%Response.Charset="utf-8"
Server.ScriptTimeout=9000
%>
<!--#include file="../inc/upload_5xsoft.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Suser")) Then Response.Redirect("../error.asp?timeout")
Dim opr,bOpen,bUpload,bRedirectToTableUpload,numUpload
Dim researchway_list
Dim conn,rs,sql,result

bOpen=True
bUpload=True
bRedirectToTableUpload=False
opr=0
sem_info=getCurrentSemester()
stu_type=Session("StuType")
researchway_list=loadResearchwayList(stu_type)

Connect conn
sql="SELECT *,dbo.getThesisStatusText(2,REVIEW_STATUS,2) AS STAT_TEXT FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("Stuid")&" ORDER BY PERIOD_ID DESC" 'AND PERIOD_ID="&sem_info(3)&" AND Valid=1"
GetRecordSetNoLock conn,rs,sql,result
sql="SELECT * FROM VIEW_STUDENT_INFO WHERE STU_ID="&Session("Stuid")
GetRecordSetNoLock conn,rs2,sql,result
tutor_duty_name=getProDutyNameOf(rs2("TUTOR_ID"))
If rs.EOF Then
	numUpload=-2
	opr=STUCLI_OPR_DETECT
	review_status=rsNone
	bUpload=False
	bRedirectToTableUpload=True
ElseIf rs("TASK_PROGRESS")<tpTbl3Passed Then
	numUpload=-2
	opr=STUCLI_OPR_DETECT
	review_status=rsNone
	bUpload=False
	bRedirectToTableUpload=True
Else
	' 评阅状态
	review_status=rs("REVIEW_STATUS")
	Select Case review_status
	Case rsNone,rsDetectThesisUploaded,rsNotAgreeDetect
		' 上传送检论文
		opr=STUCLI_OPR_DETECT
	Case rsDetected,rsReviewThesisUploaded,rsNotAgreeReview
		' 上传送审论文
		opr=STUCLI_OPR_REVIEW
	Case rsReviewEval,rsModifyThesisUploaded,rsModifyUnpassed,rsDefencePlan
		' 上传答辩论文
		opr=STUCLI_OPR_MODIFY
	Case rsModifyPassed
		opr=STUCLI_OPR_MODIFY
		bUpload=False
	Case rsInstructEval,rsFinalThesisUploaded
		' 上传定稿论文
		opr=STUCLI_OPR_FINAL
	Case Else
		bUpload=False
	End Select
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	keywords_ch=rs("KEYWORDS")
	keywords_en=rs("KEYWORDS_EN")
	researchway_name=rs("RESEARCHWAY_NAME")
	review_type=rs("REVIEW_TYPE")
	reproduct_ratio=toNumber(rs("REPRODUCTION_RATIO"))
	thesis_form=rs("THESIS_FORM")
End If
If opr<>0 Then
	bOpen=stuclient.isOpenFor(stu_type,opr)
	startdate=stuclient.getOpentime(opr,STUCLI_OPENTIME_START)
	enddate=stuclient.getOpentime(opr,STUCLI_OPENTIME_END)
	If Not bOpen Then bUpload=False
End If

If bRedirectToTableUpload Then
	CloseRs rs
	CloseConn conn
	Response.Redirect "uploadTableNew.asp"
End If

curStep=Request.QueryString("step")
Select Case curStep
Case vbNullstring ' 填写信息页面
	sql="SELECT * FROM CODE_REVIEW_TYPE WHERE LEN(THESIS_FORM)>0 AND TEACHTYPE_ID="&stu_type
	GetRecordSetNoLock conn,rs3,sql,result
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/upload.js" type="text/javascript"></script>
<script src="../scripts/uploadThesis.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<table class="tblform" width="1000"><tr><td class="summary"><p><%
	If Not bOpen Then
%><span class="tip">上传<%=arrStuOprName(opr)%>的时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，本专业上传通道已关闭或当前不在开放时间内，不能上传论文！</span><%
	ElseIf Not bUpload Then
		If numUpload=-2 Then
%><span class="tip">表格审核进度未完成，不能上传论文！</span><%
		Else
%><span class="tip">当前论文状态为【<%=rs("STAT_TEXT")%>】，不能上传论文！</span><%
		End If
	Else
%>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=arrStuOprName(opr)%></span><%
		If opr=STUCLI_OPR_DETECT Or opr=STUCLI_OPR_FINAL Then
%>&emsp;<a id="linkclaim" href="#"><img src="../images/bulb_yellow.png">查看论文提交要求</a><%
		End If
		Select Case opr
		Case STUCLI_OPR_DETECT %>
<br/>请填写以下信息，然后选择上传的文件，并点击&quot;提交&quot;按钮：<%
		Case Else %>
<br/>请选择上传的文件，然后点击&quot;提交&quot;按钮：<%
		End Select
	End If %></p></td></tr>
<tr><td align="center"><form id="fmThesis" action="?step=1" method="post" enctype="multipart/form-data">
<input type="hidden" name="uploadid" value="_student_thesisReview_uploadThesis_asp" />
<input type="hidden" name="stuid" value="<%=Session("Stuid")%>" />
<table class="tblform">
<tr><td><span class="tip">以下信息均为必填项</span></td></tr>
<tr><td align="center">
<p>论文题目：《<input type="text" name="subject_ch" size="46" maxlength="200" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="46" maxlength="200" value="<%=subject_en%>" />&nbsp;</p>
<p>作者姓名：<input type="text" name="author" size="50" value="<%=Session("Stuname")%>" readonly /></p>
<p>指导教师：<input type="text" name="tutor" size="50" value="<%=rs2("TEACHERNAME")&" "&tutor_duty_name%>" readonly /></p><%
	If stu_type=5 Then %>
<p>领域名称：<input type="text" name="speciality" size="50" value="<%=rs2("SPECIALITY_NAME")%>" readonly /></p><%
	End If %>
<p>研究方向：<%
	If stu_type=5 Or stu_type=6 Then %>
<select name="researchway_name" style="width:350px"><option value="0">请选择……</option><%
		For i=0 To UBound(researchway_list)
			option_name=researchway_list(i)
			If Left(option_name,2)="--" Then
				option_name=Mid(option_name,3)
%><option value="0" disabled><%=option_name%></option><%
			Else
%><option value="<%=option_name%>"<% If researchway_name=option_name Then %> selected<% End If %>><%=option_name%></option><%
			End If
		Next
%></select><%
	Else
%><input type="text" name="researchway_name" size="50" value="<%=researchway_name%>" /><%
	End If
%></p><%
	If opr=STUCLI_OPR_REVIEW Then	' 只在上传送审论文时显示 %>
<p>论文关键词（3-5个，用；分隔）：<input type="text" name="keywords_ch" size="46" maxlength="200" value="<%=keywords_ch%>" /></p>
<p>论文关键词（英文，3-5个，用；分隔）：<input type="text" name="keywords_en" size="39" maxlength="200" value="<%=keywords_en%>" /></p><%
	End If
%>
<p>院系名称：<input type="text" name="faculty" size="50" value="工商管理学院" readonly /></p><%
	If stu_type=5 Or stu_type=6 Then %>
<p>论文形式：<%
		If opr<>STUCLI_OPR_DETECT Then
%><input type="text" size="50" value="<%=thesis_form%>" readonly /><input type="hidden" name="thesisform" value="<%=review_type%>" /><%
		Else
%><select id="thesisform" name="thesisform" style="width:350px"><option value="0">请选择……</option><%
			Do While Not rs3.EOF
%><option value="<%=rs3("ID")%>"<% If review_type=rs3("ID") Then %> selected<% End If %>><%=rs3("THESIS_FORM")%></option><%
				rs3.MoveNext()
			Loop
%></select><%
		End If
%></p><%
	ElseIf Not rs3.EOF Then
%><select id="thesisform" name="thesisform" style="display:none"><option value="<%=rs3("ID")%>"><%=rs3("THESIS_FORM")%></option></select><%
	End If
	If opr=STUCLI_OPR_REVIEW Then %>
<p>经图书馆检测，学位论文文字复制比：<input type="text" name="reproduct_ratio" size="24" value="<%=reproduct_ratio%>%" readonly /></p><%
	End If %>
<p>论文文件：<input type="file" name="upFile" size="50" title="论文文件" /><%
	Dim jscall_checkFunc:jscall_checkFunc="checkIfWordRar"
	If bUpload Then %><br/><span class="tip"><%
		Select Case opr
		Case STUCLI_OPR_DETECT,STUCLI_OPR_MODIFY
			jscall_checkFunc="checkIfWordRar" %>
Word&nbsp;格式，文件命名为：作者姓名_学号_论文题目.doc，如“张三_201120207169_管理信息系统规划与建设研究.doc”<%
		Case STUCLI_OPR_REVIEW
			jscall_checkFunc="checkIfPDF" %>
PDF&nbsp;格式，文件命名为“盲评论文”<%
		Case STUCLI_OPR_FINAL
			jscall_checkFunc="checkIfPDF" %>
PDF&nbsp;格式，文件命名为：作者姓名_学号_论文题目.pdf<%
		End Select
%></span><br/><span class="tip">提示：超过20M请先压缩成rar文件再上传，否则上传不成功</span><%
	End If %></p><%
	If opr=STUCLI_OPR_REVIEW Then %>
<p><span class="decl">作者声明：本人确认提交的送审论文已按照华南理工大学研究生学位论文撰写规范要求排版，并已删除本人与导师的所有信息，如有问题，责任自负。</span></p><%
	End If %>
<p><input type="submit" name="btnsubmit" value="提 交"<%If Not bUpload Then %> disabled<% End If %> /></p></td></tr>
<tr><td>
<div id="divdown"><%
	If opr=STUCLI_OPR_DETECT Then
%><a href="template/fzbsmb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文文字复制比情况说明表</a><%
	Else
%><a href="template/sssqb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文分会复审意见表</a><%
	End If
%></div></td></tr></table></form></td></tr></table></center>
<div id="divclaim" class="divclaim"><%
	Select Case opr
	Case STUCLI_OPR_DETECT
%><p>送检论文提交要求：</p><p>1.电子版应为只含正文（“绪论”～“结论”部分）的Word版，须去除封面、原创性声明和使用授权书、中英文摘要、图表清单及主要符号表、目录、参考文献、附录、攻读学位期间取得的研究成果、致谢、答辩决议书、定密审批表等，电子论文命名方式为：作者姓名_学号_论文题目.doc（如“张三_201120207169_管理信息系统规划与建设研究.doc”）。电子版的格式和内容须与送审的学位论文纸质版相同。拟检测论文电子版不符合上述要求的不予检测；</p>
<p>2.拟检测论文上传成功后由导师审批，同意检测的论文由ME教育中心统一送图书馆检测；</p>
<p>3.为了保证学位论文检测的公正性和严肃性，图书馆对学位论文学术不端行为检测系统检测的结果不做任何处理，提供的《文本复制检测报告单》加盖图书馆学位论文检测章；</p>
<p>4.每篇学位论文图书馆只能检测一次，复制比要低于10%。</p><%
	Case STUCLI_OPR_REVIEW
%><p>送审论文提交要求：</p><p>提交PDF版本 ，请按照研究生院论文撰写要求排版，只需删除个人及导师信息。</p><%
	Case STUCLI_OPR_FINAL
%><p>定稿论文提交要求：</p><p>提交PDF版本 ，提交前需检查：</p>
<p>1.MBA/ME论文分类号C93、MPAcc论文分类号F23、学校代码10561、学号、论文提交日期、论文答辩日期、学位授予日期、答辩委员会成员是否填写完整（如忘记请查看学位评定意见）；</p>
<p>2.论文请勾不保密（不涉及国家机密的都不属于保密范围），原创声明一页上下作者签名、导师签名、联系电话、邮箱、地址都要填写完整，与下面的原创声明一致才是最新版本；</p>
<p>3.论文最后一页必须附学位评定意见。</p><%
	End Select %>
<p align="center"><input type="button" id="btnclaimok" value="我知道了" /></p></div>
<script type="text/javascript">
	$().ready(function() {
		$('input[name="upFile"]').change(function() {
			if(this.value.length)<%=jscall_checkFunc%>(this);
		});
		$('form').submit(function() {
			var valid=<%=jscall_checkFunc%>(this.upFile)&&checkKeywords();
			if(valid) submitUploadForm(this); else return false;
		});
		$('#btnclaimok').click(function() {
			$('#divclaim').hide();
		});
		$('#linkclaim').click(function(){
			$('#divclaim').show();
		});<%
		If Not bUpload Then %>
		$('input[name="subject_ch"],input[name="subject_en"]').attr('readOnly',true);
		$('input[name="keywords_ch"],input[name="keywords_en"]').attr('readOnly',true);
		$('a.linkAdd,a.linkRemove').attr('disabled',true);
		$('input[name="researchway_name"]'.attr('readOnly',true);
		$('input[name="upFile"]').attr('readOnly',true);
		$(':submit').attr('disabled',true);<%
		Else %>
		$(':submit').removeAttr('disabled');<%
		End If
		If opr<>STUCLI_OPR_MODIFY Then %>
		$('#divclaim').show();<%
		End If %>
	});
</script></body></html><%
	CloseRs rs3
Case 1	' 上传进程

	If Not bOpen Then
		bError=True
		errdesc="上传"&arrStuOprName(opr)&"的时间为"&FormatDateTime(startdate,1)&"至"&FormatDateTime(enddate,1)&"，本专业上传通道已关闭或当前不在开放时间内，不能上传论文！"
	ElseIf Not bUpload Then
		bError=True
		If numUpload=-2 Then
			errdesc="表格审核进度未完成，不能上传论文！"
		Else
			errdesc="当前状态为【"&rs("STAT_TEXT")&"】，不能上传论文！"
		End If
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  	CloseRs rs
  	CloseRs rs2
  	CloseConn conn
		Response.End
	End If
	
	Dim fso,Upload,file
	Dim new_subject_ch,new_subject_en
	Dim new_keywords_ch,new_keywords_en
	Dim new_researchway_name,new_review_type
	
	Set Upload=New upload_5xsoft
	new_subject_ch=Trim(Upload.Form("subject_ch"))
	new_subject_en=Trim(Upload.Form("subject_en"))
	new_keywords_ch=Trim(Upload.Form("keywords_ch"))
	new_keywords_en=Trim(Upload.Form("keywords_en"))
	new_researchway_name=Trim(Upload.Form("researchway_name"))
	new_review_type=Upload.Form("thesisform")
	If Len(new_subject_ch)=0 Then
		bError=True
		errdesc="请填写论文题目！"
	ElseIf Len(new_subject_en)=0 Then
		bError=True
		errdesc="请填写论文题目（英文）！"
	ElseIf opr=STUCLI_OPR_REVIEW And Len(new_keywords_ch)=0 Then
		bError=True
		errdesc="请填写论文关键词！"
	ElseIf opr=STUCLI_OPR_REVIEW And Len(new_keywords_en)=0 Then
		bError=True
		errdesc="请填写论文关键词（英文）！"
	ElseIf Len(new_researchway_name)=0 Or new_researchway_name="0" Then
		bError=True
		errdesc="请填写研究方向！"
	ElseIf new_review_type=0 Then
		bError=True
		errdesc="请选择论文形式！"
	End If
	If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
		Set Upload=Nothing
		CloseRs rs
  	CloseRs rs2
  	CloseConn conn
		Response.End
	End If
	Set file=Upload.File("upFile")
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	
	' 检查上传目录是否存在
	strUploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	fileExt=LCase(file.FileExt)
	If opr=STUCLI_OPR_REVIEW Or opr=STUCLI_OPR_FINAL Then
		If fileExt<>"pdf" Then
			bError=True
			errdesc="所选择的不是 PDF 文件！"
		End If
	ElseIf fileExt<>"doc" And fileExt<>"docx" And fileExt<>"rar" Then
		bError=True
		errdesc="所选择的不是 Word 文件或 RAR 压缩文件！"
'	ElseIf file.FileSize>10485760 Then
'		filesize=Round(file.FileSize/1048576,2)
'		bError=True
'		errdesc="文件大小为 "&filesize&"MB，已超出限制(10MB)！"
	End If
	If Not bError Then
		' 生成日期格式文件名
		fileid=FormatDateTime(Now(),1)&Int(Timer)
		strDestFile=fileid&"."&fileExt
		strDestPath=strUploadPath&"\"&strDestFile
		byteFileSize=file.FileSize
		' 保存
		file.SaveAs strDestPath
		
		' 关联到数据库
		sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("Stuid")&" ORDER BY PERIOD_ID DESC"
		GetRecordSet conn,rs3,sql,result
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		rs3("RESEARCHWAY_NAME")=new_researchway_name
		Select Case opr
		Case STUCLI_OPR_DETECT ' 送检论文
			rs3("THESIS_FILE")=strDestFile
			rs3("REVIEW_TYPE")=new_review_type
			If review_status<rsDetectThesisUploaded Then rs3("REVIEW_STATUS")=rsDetectThesisUploaded
			rs3("DETECT_APP_EVAL")=Null
		Case STUCLI_OPR_REVIEW ' 送审论文
			rs3("THESIS_FILE2")=strDestFile
			rs3("KEYWORDS")=new_keywords_ch
			rs3("KEYWORDS_EN")=new_keywords_en
			If review_status<rsReviewThesisUploaded Then rs3("REVIEW_STATUS")=rsReviewThesisUploaded
			rs3("REVIEW_APP_EVAL")=Null
		Case STUCLI_OPR_MODIFY ' 答辩论文
			rs3("THESIS_FILE3")=strDestFile
			If review_status<rsModifyThesisUploaded Then rs3("REVIEW_STATUS")=rsModifyThesisUploaded
			rs3("TUTOR_MODIFY_EVAL")=Null
		Case STUCLI_OPR_FINAL ' 定稿论文
			rs3("THESIS_FILE4")=strDestFile
			If review_status<rsFinalThesisUploaded Then rs3("REVIEW_STATUS")=rsFinalThesisUploaded
		End Select
		rs3.Update()
		CloseRs rs3
		If opr<>STUCLI_OPR_FINAL Then
			' 向导师发送审核通知邮件
			'sendEmailToTutor arrStuOprName(opr)
		End If
		Dim logtxt
		logtxt="学生["&Session("Stuname")&"]上传["&arrStuOprName(opr)&"]。"
		WriteLogForReviewSystem logtxt
	End If
	Set fso=Nothing
	Set file=Nothing
	Set Upload=Nothing
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传专业硕士论文</title>
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
</head>
<body bgcolor="ghostwhite">
<center><br/><b>上传专业硕士论文</b><br/><br/><%
	If Not bError Then %>
<form id="fmFinish" action="default.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestFile%>" />
<p><%=byteFileSize%> 字节已上传，正在关联数据...</p></form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%>
<script type="text/javascript">alert("<%=errdesc%>");history.go(-1);</script><%
	End If
%></center></body></html><%
End Select
CloseRs rs2
CloseRs rs
CloseConn conn
%>