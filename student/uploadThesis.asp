<%'Response.Charset="utf-8"
Server.ScriptTimeout=9000
%><!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/mail.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim opr,bOpen,bUpload,bRedirectToTableUpload,bTableReady
Dim researchway_list
Dim conn,rs,sql,result

bOpen=True
bUpload=True
bRedirectToTableUpload=False
bTableReady=True
opr=0
sem_info=getCurrentSemester()
stu_type=Session("StuType")
researchway_list=loadResearchwayList(stu_type)

Connect conn
sql="SELECT *,dbo.getThesisStatusText(2,REVIEW_STATUS,2) AS STAT_TEXT FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("StuId")&" ORDER BY PERIOD_ID DESC" 'AND PERIOD_ID="&sem_info(3)&" AND Valid=1"
GetRecordSetNoLock conn,rs,sql,result
sql="SELECT * FROM VIEW_STUDENT_INFO WHERE STU_ID="&Session("StuId")
GetRecordSetNoLock conn,rs2,sql,result
tutor_duty_name=getProDutyNameOf(rs2("TUTOR_ID"))
If rs.EOF Then
	bTableReady=False
	opr=STUCLI_OPR_DETECT_REVIEW
	review_status=rsNone
	bUpload=False
	bRedirectToTableUpload=True
ElseIf rs("TASK_PROGRESS").Value<tpTbl3Passed Then
	bTableReady=False
	opr=STUCLI_OPR_DETECT_REVIEW
	review_status=rsNone
	bUpload=False
	bRedirectToTableUpload=True
Else
	' 评阅状态
	review_status=rs("REVIEW_STATUS").Value
	Select Case review_status
	Case rsNone,rsDetectThesisUploaded,rsNotAgreeDetect,rsDetectUnpassed
		' 上传送检论文和送审论文
		opr=STUCLI_OPR_DETECT_REVIEW
	Case rsNotAgreeReview
		' 上传送审论文
		opr=STUCLI_OPR_REVIEW
	Case rsReviewEval,rsModifyThesisUploaded
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
	subject_ch=rs("THESIS_SUBJECT").Value
	subject_en=rs("THESIS_SUBJECT_EN").Value
	keywords_ch=rs("KEYWORDS").Value
	keywords_en=rs("KEYWORDS_EN").Value
	researchway_name=rs("RESEARCHWAY_NAME").Value
	review_type=rs("REVIEW_TYPE").Value
	reproduct_ratio=toNumber(rs("REPRODUCTION_RATIO").Value)
	thesis_form=rs("THESIS_FORM").Value
End If
If opr<>0 Then
	bOpen=stuclient.isOpenFor(stu_type,opr)
	startdate=gstuclient.getOpentime(opr,STUCLI_OPENTIME_START)
	enddate=stuclient.getOpentime(opr,STUCLI_OPENTIME_END)
	If Not bOpen Then bUpload=False
	uploadTypename=arrStuOprName(opr)
	If opr=STUCLI_OPR_DETECT_REVIEW Then
		n=Sgn(rs("DETECT_COUNT").Value)
		subtype=Array("一次","二次")(n)
		uploadTypename=subtype&uploadTypename
	End If
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
<table class="tblform" width="1000"><tr><td class="summary"><p align="center"><%
	If Not bOpen Then
%><span class="tip">上传<%=arrStuOprName(opr)%>的时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，本专业上传通道已关闭或当前不在开放时间内，不能上传论文！</span><%
	ElseIf Not bUpload Then
		If Not bTableReady Then
%><span class="tip">表格审核进度未完成，不能上传论文！</span><%
		Else
%><span class="tip">当前论文状态为【<%=rs("STAT_TEXT").Value%>】，不能上传论文！</span><%
		End If
	Else
%>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=uploadTypename%></span><%
		If opr=STUCLI_OPR_DETECT_REVIEW Or opr=STUCLI_OPR_REVIEW Or opr=STUCLI_OPR_FINAL Then
%>&emsp;<a id="linkclaim" href="#"><img src="../images/bulb_yellow.png">查看论文提交要求</a><%
		End If
		Select Case opr
		Case STUCLI_OPR_DETECT_REVIEW %>
<br/>请填写以下信息，然后选择上传的文件，并点击&quot;提交&quot;按钮：<%
		Case Else %>
<br/>请选择上传的文件，然后点击&quot;提交&quot;按钮：<%
		End Select
	End If %></p></td></tr>
<tr><td align="center"><form id="fmThesis" action="?step=1" method="post" enctype="multipart/form-data">
<input type="hidden" name="uploadid" value="_student_thesisReview_uploadThesis_asp" />
<input type="hidden" name="StuId" value="<%=Session("StuId")%>" />
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
	If opr=STUCLI_OPR_DETECT_REVIEW OR opr=STUCLI_OPR_REVIEW Then	' 只在上传送审论文时显示 %>
<p>论文关键词（3-5个，用；分隔）：<input type="text" name="keywords_ch" size="46" maxlength="200" value="<%=keywords_ch%>" /></p>
<p>论文关键词（英文，3-5个，用；分隔）：<input type="text" name="keywords_en" size="39" maxlength="200" value="<%=keywords_en%>" /></p><%
	End If
%>
<p>院系名称：<input type="text" name="faculty" size="50" value="工商管理学院" readonly /></p><%
	If stu_type=5 Or stu_type=6 Then %>
<p>论文形式：<%
		If opr<>STUCLI_OPR_DETECT_REVIEW Then
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
	End If
	
	Dim callbackValidate:callbackValidate="checkIfWordRar"
	If opr=STUCLI_OPR_DETECT_REVIEW Then
		callbackValidate="checkIfDetectReview"
%><p>送检论文文件：<input type="file" name="detectFile" size="50" title="送检论文文件" /><%
		If bUpload Then
%><br/><span class="tip">Word&nbsp;格式，文件命名为：作者姓名_学号_论文题目.doc，如“张三_201120207169_管理信息系统规划与建设研究.doc”</span><%
		End If
%></p><p>送审论文文件：<input type="file" name="reviewFile" size="50" title="送审论文文件" /><%
		If bUpload Then
%><br/><span class="tip">PDF&nbsp;格式，文件命名为“盲评论文”</span><%
		End If
%></p><%
	Else
%><p>论文文件：<input type="file" name="upFile" size="50" title="论文文件" /><%
		If bUpload Then %><br/><span class="tip"><%
			Select Case opr
			Case STUCLI_OPR_REVIEW
				callbackValidate="checkIfPdfRar" %>
PDF&nbsp;格式，文件命名为“盲评论文”<%
			Case STUCLI_OPR_MODIFY
				callbackValidate="checkIfWordRar" %>
Word&nbsp;格式，文件命名为：作者姓名_学号_论文题目.doc，如“张三_201120207169_管理信息系统规划与建设研究.doc”<%
			Case STUCLI_OPR_FINAL
				callbackValidate="checkIfPdfRar" %>
PDF&nbsp;格式，文件命名为：作者姓名_学号_论文题目.pdf<%
			End Select
%></span><%
		End If
%></p><%
	End If
%><span class="tip">提示：超过20M请先压缩成rar文件再上传，否则上传不成功</span><%
	If opr=STUCLI_OPR_DETECT_REVIEW Or opr=STUCLI_OPR_REVIEW Then %>
<p class="decl">作者承诺：<br/>
1．该学位论文为公开学位论文，其中不涉及国家秘密项目和其它不宜公开的内容，否则将由本人承担因学位论文涉密造成的损失和相关的法律责任；<br/>
2．该学位论文是本人在导师的指导下独立进行研究所取得的研究成果，不存在学术不端行为。</p><%
	End If %>
<p><input type="submit" name="btnsubmit" value="提 交"<%If Not bUpload Then %> disabled<% End If %> /></p></td></tr>
<tr><td>
<div id="divdown" style="display: none">
<p><a href="template/fzbsmb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文文字复制比情况说明表</a></p>
<p><a href="template/sssqb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文分会复审意见表</a></p>
</div></td></tr></table></form></td></tr></table></center>
<div id="divclaim" class="divclaim"><%
	Select Case opr
	Case STUCLI_OPR_DETECT_REVIEW
%><p>检测论文与送审论文要求是同一篇论文，按不同格式要求提交。</p>
<h1>送检论文提交要求：</h1><p>电子版应为只含正文（“绪论”～“结论”部分）的Word版，须去除封面、原创性声明和使用授权书、中英文摘要、图表清单及主要符号表、目录、参考文献、附录、攻读学位期间取得的研究成果、致谢、答辩决议书、定密审批表等，电子论文命名方式为：作者姓名_学号_论文题目.doc（如“张三_201120207169_管理信息系统规划与建设研究.doc”）。</p>
<h1>送审论文提交要求：</h1><p>PDF版本，文件命名为“盲评论文”，按照华南理工大学研究生学位论文撰写规范要求排版，学位论文中涉及个人学号、姓名及导师姓名的部分全部留空，致谢部分不出现任何人的姓名。</p>
<p>检测论文与送审论文由导师审核通过，由MBA/MPAcc/ME教育中心统一报图书馆进行检测。</p>
<h1>检测结果处理：</h1>
<ol><li>复制比低于10%的学员，系统会自动匹配进行论文盲审。</li>
<li>复制比高于10%的学员，导师应对学位论文的学术规范性进行严格把关，学员在导师的督促下对论文中存在的学术不规范部分进行修改。修改后的论文，如导师同意再次进行送审，学员需登录系统再次提交检测和送审论文，<span class="stress">由中心统一进行二次查重，二次查重所产生的费用由学员本人缴纳。</span></li>
<li>导师若发现学位论文存在学术不端行为的情况将上报学院，学院根据《华南理工大学研究生学术道德规范及学术不端行为处理办法》提出初步处理意见并上报学位办，由学校作出处理。</li></ol>
<h1>评审结果处理：</h1><p>评审结果分为同意答辩、适当修改后可答辩、须做重大修改后方可答辩、不同意答辩四种情况。具体的参照《论文撰写规范及流程》执行</p><%
	Case STUCLI_OPR_REVIEW
%><p>送审论文提交要求：</p><p>提交PDF版本，请按照研究生院论文撰写要求排版，只需删除个人及导师信息。</p><%
	Case STUCLI_OPR_FINAL
%><p>定稿论文提交要求：</p><p>提交PDF版本，提交前需检查：</p>
<p>1.MBA/ME论文分类号C93、MPAcc论文分类号F23、学校代码10561、学号、论文提交日期、论文答辩日期、学位授予日期、答辩委员会成员是否填写完整（如忘记请查看学位评定意见）；</p>
<p>2.论文请勾不保密（不涉及国家机密的都不属于保密范围），原创声明一页上下作者签名、导师签名、联系电话、邮箱、地址都要填写完整，与下面的原创声明一致才是最新版本；</p>
<p>3.论文最后一页必须附学位评定意见。</p><%
	End Select %>
<p align="center"><input type="button" id="btnclaimok" value="我知道了" /></p></div>
<script type="text/javascript">
	$().ready(function() {
		$(':file').change(function() {
			if(this.value.length)<%=callbackValidate%>(this,$(this).index());
		});
		$('form').submit(function() {
			var valid=<%=callbackValidate%>($(':file'))&&checkKeywords();
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
		$(':file').attr('readOnly',true);
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
		If Not bTableReady Then
			errdesc="表格审核进度未完成，不能上传论文！"
		Else
			errdesc="当前状态为【"&rs("STAT_TEXT").Value&"】，不能上传论文！"
		End If
	End If
	If bError Then
		%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  	CloseRs rs
  	CloseRs rs2
  	CloseConn conn
		Response.End
	End If
	
	Dim fso,Upload,file,file2
	Dim new_subject_ch,new_subject_en
	Dim new_keywords_ch,new_keywords_en
	Dim new_researchway_name,new_review_type
	Dim sqlDetect
	
	Set Upload=New ExtendedRequest
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
	
	Set fso=Server.CreateObject("Scripting.FileSystemObject")
	' 检查上传目录是否存在
	uploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(uploadPath) Then fso.CreateFolder(uploadPath)
	Select Case opr
	Case STUCLI_OPR_DETECT_REVIEW
		Set file=Upload.File("detectFile")
		Set file2=Upload.File("reviewFile")
		fileExt=LCase(file.FileExt)
		fileExt2=LCase(file2.FileExt)
		If fileExt<>"doc" And fileExt<>"docx" And fileExt<>"rar" Then
			bError=True
			errdesc="所选择的不是 Word 文件或 RAR 压缩文件！"
		End If
		If fileExt2<>"pdf" And fileExt2<>"rar" Then
			bError=True
			errdesc="所选择的不是 PDF 文件或 RAR 压缩文件！"
		End If
	Case STUCLI_OPR_REVIEW,STUCLI_OPR_FINAL
		Set file=Upload.File("upFile")
		fileExt=LCase(file.FileExt)
		If fileExt<>"pdf" Then
			bError=True
			errdesc="所选择的不是 PDF 文件或 RAR 压缩文件！"
		End If
	Case Else
		Set file=Upload.File("upFile")
		fileExt=LCase(file.FileExt)
		If fileExt<>"doc" And fileExt<>"docx" And fileExt<>"rar" Then
			bError=True
			errdesc="所选择的不是 Word 文件或 RAR 压缩文件！"
		End If
	End Select
	If Not bError Then
		' 生成日期格式文件名
		fileid=FormatDateTime(Now(),1)&Int(Timer)
		destFile=fileid&"."&fileExt
		destPath=uploadPath&"\"&destFile
		fileSize=file.FileSize
		' 保存
		file.SaveAs destPath
		
		' 关联到数据库
		sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE STU_ID="&Session("StuId")&" ORDER BY PERIOD_ID DESC"
		GetRecordSet conn,rs3,sql,result
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		rs3("RESEARCHWAY_NAME")=new_researchway_name
		Select Case opr
		Case STUCLI_OPR_DETECT_REVIEW ' 送检论文和送审论文
			destFile2=fileid&"1."&fileExt2
			destPath2=uploadPath&"\"&destFile2
			file2.SaveAs destPath2
			
			If review_status=rsDetectThesisUploaded Then
				sqlDetect="EXEC spDeleteDetectResult "&rs("ID").Value&","&toSqlString(rs3("THESIS_FILE").Value)&";"
			End If
			sqlDetect=sqlDetect&"EXEC spAddDetectResult "&rs("ID").Value&","&toSqlString(destFile)&",NULL,NULL,NULL;"
			rs3("THESIS_FILE")=destFile
			rs3("THESIS_FILE2")=destFile2
			rs3("KEYWORDS")=new_keywords_ch
			rs3("KEYWORDS_EN")=new_keywords_en
			rs3("REVIEW_TYPE")=new_review_type
			rs3("REVIEW_STATUS")=rsDetectThesisUploaded
			If rs("DETECT_COUNT").Value=0 Then
				rs3("DETECT_APP_EVAL")=Null
				rs3("REVIEW_APP_EVAL")=Null
			End If
		Case STUCLI_OPR_REVIEW ' 送审论文
			rs3("THESIS_FILE2")=destFile
			rs3("KEYWORDS")=new_keywords_ch
			rs3("KEYWORDS_EN")=new_keywords_en
			rs3("REVIEW_STATUS")=rsRedetectPassed
			rs3("REVIEW_APP_EVAL")=Null
		Case STUCLI_OPR_MODIFY ' 答辩论文
			rs3("THESIS_FILE3")=destFile
			rs3("REVIEW_STATUS")=rsModifyThesisUploaded
			rs3("TUTOR_MODIFY_EVAL")=Null
		Case STUCLI_OPR_FINAL ' 定稿论文
			rs3("THESIS_FILE4")=destFile
			rs3("REVIEW_STATUS")=rsFinalThesisUploaded
		End Select
		rs3.Update()
		CloseRs rs3
		
		If Len(sqlDetect) Then
			conn.Execute sqlDetect
		End If
		' If opr<>STUCLI_OPR_FINAL Then
			' 向导师发送审核通知邮件
			' sendEmailToTutor arrStuOprName(opr)
		' End If
		Dim logtxt
		logtxt="学生["&Session("StuName")&"]上传["&arrStuOprName(opr)&"]。"
		WriteLog logtxt
	End If
	Set fso=Nothing
	Set file=Nothing
	Set file2=Nothing
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
<form id="fmFinish" action="home.asp" method="post"><input type="hidden" name="filename" value="<%=destFile%>" /></form>
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