<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim section_id,time_flag,uploadable,redirect_to_tbl_upload,table_ready
Dim researchway_list
Dim conn,rs,sql,count

uploadable=False
redirect_to_tbl_upload=False
table_ready=True
section_id=0
stu_type=Session("StuType")

ConnectDb conn
sql="SELECT *,dbo.getThesisStatusText(2,REVIEW_STATUS,2) AS STAT_TEXT FROM ViewDissertations WHERE STU_ID="&Session("StuId")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
sql="SELECT * FROM ViewStudentInfo WHERE STU_ID="&Session("StuId")
GetRecordSetNoLock conn,rsStu,sql,count
tutor_duty_name=getProDutyNameOf(rsStu("TUTOR_ID"))
If rs.EOF Then
	table_ready=False
	section_id=sectionUploadDetectReview
	review_status=rsNone
	redirect_to_tbl_upload=True
ElseIf rs("TASK_PROGRESS")<tpTbl3Passed Then
	table_ready=False
	section_id=sectionUploadDetectReview
	review_status=rsNone
	redirect_to_tbl_upload=True
Else
	' 评阅状态
	review_status=rs("REVIEW_STATUS")
	Select Case review_status
	Case rsNone,rsDetectPaperUploaded,rsRefusedDetect,rsDetectUnpassed
		' 上传送检论文和送审论文
		section_id=sectionUploadDetectReview
	Case rsReviewEval,rsDefencePaperUploaded,rsRefusedDefence
		' 上传答辩论文
		section_id=sectionUploadDefence
	Case rsDefenceEval,rsInstructReviewPaperUploaded,rsRefusedInstructReview
		' 上传教指委论文
		section_id=sectionUploadInstructReview
	Case rsInstructEval,rsFinalPaperUploaded
		' 上传定稿论文
		section_id=sectionUploadFinal
	End Select
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	keywords_ch=rs("KEYWORDS")
	keywords_en=rs("KEYWORDS_EN")
	sub_research_field=rs("RESEARCHWAY_NAME")
	review_type=rs("REVIEW_TYPE")
	reproduct_ratio=toNumericString(rs("REPRODUCTION_RATIO"))
	thesis_form=rs("THESIS_FORM")
End If
If section_id<>0 Then
	If rs.EOF Then
		uploadable=True
	ElseIf Not isActivityOpen(rs("ActivityId")) Then
		time_flag=-3
	Else
		Set current_section=getSectionInfo(rs("ActivityId"), stu_type, section_id)
		time_flag=compareNowWithSectionTime(current_section)
		uploadable=time_flag=0
		upload_stuff_name=arrStuOprName(section_id)
		If section_id=sectionUploadDetectReview Then
			If rs.EOF Then n=0 Else n=Sgn(rs("DETECT_COUNT"))
			subtype=Array("一次","二次")(n)
			upload_stuff_name=subtype&upload_stuff_name
		End If
	End If
End If

If redirect_to_tbl_upload Then
	CloseRs rs
	CloseConn conn
	Response.Redirect "uploadTable.asp"
End If

step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 填写信息页面
	sql="SELECT * FROM ReviewTypes WHERE LEN(THESIS_FORM)>0 AND TEACHTYPE_ID="&stu_type
	GetRecordSetNoLock conn,rs3,sql,count
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传论文</title>
<% useStylesheet "student" %>
<% useScript "jquery", "upload", "uploadPaper" %>
</head>
<body>
<table class="form" width="1000" align="center"><tr><td class="summary"><p><%
	If Not uploadable Then
		If time_flag=-3 Then
%><span class="tip">当前评阅活动【<%=rs("ActivityName")%>】已关闭，不能上传论文！</span><%
		ElseIf time_flag=-2 Then
%><span class="tip">【<%=current_section("Name")%>】环节已关闭，不能上传论文！</span><%
		ElseIf time_flag<>0 Then
%><span class="tip">【<%=current_section("Name")%>】环节开放时间为<%=toDateTime(current_section("StartTime"),1)%>至<%=toDateTime(current_section("EndTime"),1)%>，当前不在开放时间内，不能上传论文！</span><%
		ElseIf Not table_ready Then
%><span class="tip">表格审核进度未完成，不能上传论文！</span><%
		Else
%><span class="tip">当前论文状态为【<%=rs("STAT_TEXT")%>】，不能上传论文！</span><%
		End If
	Else
%>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=upload_stuff_name%></span><%
		If section_id=sectionUploadDetectReview Or section_id=sectionUploadReview Or section_id=sectionUploadFinal Then
%>&emsp;<a id="linkclaim" href="#"><img src="../images/bulb_yellow.png">查看论文提交要求</a><%
		End If
		Select Case section_id
		Case sectionUploadDetectReview %>
<br/>请填写以下信息，然后选择上传的文件，并点击&quot;提交&quot;按钮：<%
		Case Else %>
<br/>请选择上传的文件，然后点击&quot;提交&quot;按钮：<%
		End Select
	End If %></p></td></tr>
<tr><td><form id="fmDissertation" action="?step=1" method="post" enctype="multipart/form-data">
<input type="hidden" name="upload_id" value="_student_thesisReview_uploadThesis_asp" />
<input type="hidden" name="stu_id" value="<%=Session("StuId")%>" />
<table class="form" width="800">
<tr><td align="center"><span class="tip">以下信息均为必填项</span></td></tr>
<tr><td>
<p>论文题目：《<input type="text" name="subject_ch" size="70" maxlength="200" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="73" maxlength="200" value="<%=subject_en%>" />&nbsp;</p>
<p>作者姓名：<input type="text" name="author" size="50" value="<%=Session("Stuname")%>" readonly /></p>
<p>指导教师：<input type="text" name="tutor" size="50" value="<%=rsStu("TEACHERNAME")&" "&tutor_duty_name%>" readonly /></p><%
	If stu_type=5 Then %>
<p>领域名称：<input type="text" name="speciality" size="50" value="<%=rsStu("SPECIALITY_NAME")%>" readonly /></p><%
	End If %>
<p>研究方向：
<select name="sub_research_field_select" style="width:350px"></select><span class="tip">请务必认真选择，系统根据研究方向匹配专家</span>
<div class="custom-sub-research-field"><input type="text" name="custom_sub_research_field" size="50" placeholder="请输入…" value="<%=sub_research_field%>" /></div>
<input type="hidden" name="sub_research_field" value="<%=sub_research_field%>" /></p><%
	If section_id=sectionUploadDetectReview Or section_id=sectionUploadReview Then	' 只在上传送审论文时显示 %>
<p>论文关键词（3-5个，用；分隔）：<input type="text" name="keywords_ch" size="46" maxlength="200" value="<%=keywords_ch%>" /></p>
<p>论文关键词（英文，3-5个，用；分隔）：<input type="text" name="keywords_en" size="39" maxlength="200" value="<%=keywords_en%>" /></p><%
	End If
%>
<p>院系名称：<input type="text" name="faculty" size="50" value="工商管理学院" readonly /></p><%
	If stu_type=5 Or stu_type=6 Then %>
<p>论文形式：<%
		If section_id<>sectionUploadDetectReview Then
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

	Dim callbackValidate:callbackValidate="checkIfWordRar"
	If section_id=sectionUploadDetectReview Then
		callbackValidate="checkIfDetectReview"
%><p>送检论文文件：<input type="file" name="detectFile" size="50" title="送检论文文件" /><%
		If uploadable Then
%><span class="tip">Word&nbsp;格式</span><%
		End If
%></p><p>送审论文文件：<input type="file" name="reviewFile" size="50" title="送审论文文件" /><%
		If uploadable Then
%><span class="tip">PDF&nbsp;格式</span><%
		End If
%></p><%
	Else
%><p>论文文件：<input type="file" name="upFile" size="50" title="论文文件" /><%
		If uploadable Then %><span class="tip"><%
			Select Case section_id
			Case sectionUploadReview
				callbackValidate="checkIfPdfRar" %>
PDF&nbsp;格式<%
			Case sectionUploadDefence
				callbackValidate="checkIfWordRar" %>
Word&nbsp;格式<%
			Case sectionUploadInstructReview
				callbackValidate="checkIfPdfRar" %>
PDF&nbsp;格式<%
			Case sectionUploadFinal
				callbackValidate="checkIfPdfRar" %>
PDF&nbsp;格式<%
			End Select
%></span><%
		End If
%></p><%
	End If
%><p align="center"><span class="tip">提示：超过20M请先压缩成rar文件再上传，否则上传不成功</span></p><%
	If section_id=sectionUploadDetectReview Or section_id=sectionUploadReview Then %>
<p class="decl">作者承诺：<br/>
1．该学位论文为公开学位论文，其中不涉及国家秘密项目和其它不宜公开的内容，否则将由本人承担因学位论文涉密造成的损失和相关的法律责任；<br/>
2．该学位论文是本人在导师的指导下独立进行研究所取得的研究成果，不存在学术不端行为。</p><%
	End If %>
<p align="center"><input type="submit" name="btnsubmit" value="提 交"<%If Not uploadable Then %> disabled<% End If %> /></p></td></tr>
<tr><td>
<div id="divdown" style="display: none">
<p><a href="template/fzbsmb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文文字复制比情况说明表</a></p>
<p><a href="template/sssqb.doc" target="_blank"><img src="../images/down.png" />下载硕士学位论文分会复审意见表</a></p>
</div></td></tr></table></form></td></tr></table></center>
<div id="claim" class="claim"><%
	Select Case section_id
	Case sectionUploadDetectReview
%><p>检测论文与送审论文要求是同一篇论文，按不同格式要求提交。</p>
<h1>送检论文提交要求：</h1><p>电子版应为只含正文（“绪论”～“结论”部分）的Word版，须去除封面、原创性声明和使用授权书、中英文摘要、图表清单及主要符号表、目录、参考文献、附录、攻读学位期间取得的研究成果、致谢、答辩决议书、定密审批表等，电子论文命名方式为：作者姓名_学号_论文题目.doc（如“张三_201120207169_管理信息系统规划与建设研究.doc”）。</p>
<h1>送审论文提交要求：</h1><p>PDF版本，文件命名为“盲评论文”，按照华南理工大学研究生学位论文撰写规范要求排版，学位论文中涉及个人学号、姓名及导师姓名的部分全部留空，致谢部分不出现任何人的姓名。</p>
<p>检测论文与送审论文由导师审核通过，由MBA/MPAcc/ME/EMBA教育中心统一报图书馆进行检测。</p>
<h1>检测结果处理：</h1>
<ol><li>复制比低于10%的学员，系统会自动匹配进行论文盲审。</li>
<li>复制比高于10%的学员，导师应对学位论文的学术规范性进行严格把关，学员在导师的督促下对论文中存在的学术不规范部分进行修改。修改后的论文，如导师同意再次进行送审，学员需登录系统再次提交检测和送审论文，<span class="stress">由中心统一进行二次查重，二次查重所产生的费用由学员本人缴纳。</span></li>
<li>导师若发现学位论文存在学术不端行为的情况将上报学院，学院根据《华南理工大学研究生学术道德规范及学术不端行为处理办法》提出初步处理意见并上报学位办，由学校作出处理。</li></ol>
<h1>评审结果处理：</h1><p>评审结果分为同意答辩、适当修改后可答辩、须做重大修改后方可答辩、不同意答辩四种情况。具体的参照《论文撰写规范及流程》执行</p><%
	Case sectionUploadReview
%><p>送审论文提交要求：</p><p>提交PDF版本，请按照研究生院论文撰写要求排版，只需删除个人及导师信息。</p><%
	Case sectionUploadInstructReview
%><p>教指委盲评论文提交要求：</p><ol><li>必须与所提交的纸质版论文完全一致；</li>
<li>按研究生院论文撰写规范排版；</li>
<li>学位论文中涉及个人学号、姓名及导师姓名的部分全部留空；</li>
<li>在读期间所取得的学术成果列表中只列刊物名称和卷、期号，如无此表格留空；</li>
<li>致谢部分不出现任何人的姓名。</li></ol><%
	Case sectionUploadFinal
%><p>定稿论文提交要求：</p><p>提交PDF版本，提交前需检查：</p>
<p>1.MBA/ME论文分类号C93、MPAcc论文分类号F23、学校代码10561、学号、论文提交日期、论文答辩日期、学位授予日期、答辩委员会成员是否填写完整（如忘记请查看学位评定意见）；</p>
<p>2.论文请勾不保密（不涉及国家机密的都不属于保密范围），原创声明一页上下作者签名、导师签名、联系电话、邮箱、地址都要填写完整，与下面的原创声明一致才是最新版本；</p>
<p>3.论文最后一页必须附学位评定意见。</p><%
	End Select %>
<p align="center"><input type="button" id="btnclaimok" value="我知道了" /></p></div>
<script type="text/javascript">
	$().ready(function() {
		$('select[name="sub_research_field_select"]').change(function() {
			$('input[name="sub_research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
			$custom_field=$('input[name="custom_sub_research_field"]');
			if(this.value=='-1') {
				$custom_field.parents('div').eq(0).show();
				$custom_field.focus();
			} else {
				$custom_field.parents('div').eq(0).hide();
			}
		});
		$(':file').change(function() {
			if(this.value.length)<%=callbackValidate%>($(this),$(':file').index(this));
		});
		$('form').submit(function() {
			var valid=<%=callbackValidate%>($(':file'));
			if(valid) submitUploadForm(this); else return false;
		});
		$('#btnclaimok').click(function() {
			$('#claim').hide();
		});
		$('#linkclaim').click(function(){
			$('#claim').show();
		});
		initAllSubResearchFieldSelectBox($('select[name="sub_research_field_select"]'),<%=stu_type%>,'<%=sub_research_field%>');
		<%
		If Not uploadable Then %>
		$('input[name="subject_ch"],input[name="subject_en"]').attr('readOnly',true);
		$('input[name="keywords_ch"],input[name="keywords_en"]').attr('readOnly',true);
		$('a.linkAdd,a.linkRemove').attr('disabled',true);
		$('input[name="sub_research_field"]').attr('readOnly',true);
		$(':file').attr('readOnly',true);
		$(':submit').attr('disabled',true);<%
		Else %>
		$(':submit').removeAttr('disabled');<%
			If section_id<>sectionUploadDefence Then %>
			$('#claim').show();<%
			End If
		End If%>
	});
</script></body></html><%
	CloseRs rs3
Case 1	' 上传进程

	If time_flag=-3 Then
		bError=True
		errMsg=Format("当前评阅活动【{0}】已关闭，不能提交表格！", rs("ActivityName"))
	ElseIf time_flag=-2 Then
		bError=True
		errMsg=Format("【{0}】环节已关闭，不能上传论文！",current_section("Name"))
	ElseIf time_flag<>0 Then
		bError=True
		errMsg=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能上传论文！",_
			current_section("Name"),_
			toDateTime(current_section("StartTime"),1),_
			toDateTime(current_section("EndTime"),1))
	ElseIf Not uploadable Then
		bError=True
		If Not table_ready Then
			errMsg="表格审核进度未完成，不能上传论文！"
		Else
			errMsg="当前状态为【"&rs("STAT_TEXT")&"】，不能上传论文！"
		End If
	End If
	If bError Then
		CloseRs rs
		CloseRs rsStu
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If

	Dim fso,Upload,file,file2
	Dim new_subject_ch,new_subject_en
	Dim new_keywords_ch,new_keywords_en
	Dim new_sub_research_field_id,new_sub_research_field,custom_sub_research_field
	Dim new_review_type
	Dim sqlDetect

	Set Upload=New ExtendedRequest
	new_subject_ch=Trim(Upload.Form("subject_ch"))
	new_subject_en=Trim(Upload.Form("subject_en"))
	new_keywords_ch=Trim(Upload.Form("keywords_ch"))
	new_keywords_en=Trim(Upload.Form("keywords_en"))
	new_sub_research_field_id=Upload.Form("sub_research_field_select")
	new_sub_research_field=Upload.Form("sub_research_field")
	custom_sub_research_field=Trim(Upload.Form("custom_sub_research_field"))
	new_review_type=Upload.Form("thesisform")
	If Len(new_subject_ch)=0 Then
		bError=True
		errMsg="请填写论文题目！"
	ElseIf Len(new_subject_en)=0 Then
		bError=True
		errMsg="请填写论文题目（英文）！"
	ElseIf (section_id=sectionUploadDetectReview Or section_id=sectionUploadReview) And Len(new_keywords_ch)=0 Then
		bError=True
		errMsg="请填写论文关键词！"
	ElseIf (section_id=sectionUploadDetectReview Or section_id=sectionUploadReview) And Len(new_keywords_en)=0 Then
		bError=True
		errMsg="请填写论文关键词（英文）！"
	ElseIf Len(new_sub_research_field_id)=0 Then
		bError=True
		errMsg="请选择研究方向！"
	ElseIf new_sub_research_field_id="-1" And Len(custom_sub_research_field)=0 Then
		bError=True
		errMsg="请填写研究方向！"
	ElseIf new_review_type=0 Then
		bError=True
		errMsg="请选择论文形式！"
	End If
	If bError Then
		Set Upload=Nothing
		CloseRs rs
		CloseRs rsStu
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If

	Set fso=CreateFSO()
	' 检查上传目录是否存在
	uploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(uploadPath) Then fso.CreateFolder(uploadPath)
	Select Case section_id
	Case sectionUploadDetectReview
		Set file=Upload.File("detectFile")
		Set file2=Upload.File("reviewFile")
		file_ext=LCase(file.FileExt)
		fileExt2=LCase(file2.FileExt)
		If file_ext<>"doc" And file_ext<>"docx" And file_ext<>"rar" Then
			bError=True
			errMsg="所选择的不是 Word 文件或 RAR 压缩文件！"
		End If
		If fileExt2<>"pdf" And fileExt2<>"rar" Then
			bError=True
			errMsg="所选择的不是 PDF 文件或 RAR 压缩文件！"
		End If
	Case sectionUploadReview,sectionUploadInstructReview,sectionUploadFinal
		Set file=Upload.File("upFile")
		file_ext=LCase(file.FileExt)
		If file_ext<>"pdf" And file_ext<>"rar" Then
			bError=True
			errMsg="所选择的不是 PDF 文件或 RAR 压缩文件！"
		End If
	Case Else
		Set file=Upload.File("upFile")
		file_ext=LCase(file.FileExt)
		If file_ext<>"doc" And file_ext<>"docx" And file_ext<>"rar" Then
			bError=True
			errMsg="所选择的不是 Word 文件或 RAR 压缩文件！"
		End If
	End Select
	If Not bError Then
		' 生成日期格式文件名
		fileid=timestamp()
		destFile=fileid&"."&file_ext
		destPath=uploadPath&"\"&destFile
		fileSize=file.FileSize
		' 保存
		file.SaveAs destPath

		' 关联到数据库
		sql="SELECT * FROM Dissertations WHERE STU_ID="&Session("StuId")&" ORDER BY ActivityId DESC"
		GetRecordSet conn,rs3,sql,count
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		If new_sub_research_field_id="-1" Then new_sub_research_field=custom_sub_research_field
		rs3("RESEARCHWAY_NAME")=new_sub_research_field
		Select Case section_id
		Case sectionUploadDetectReview ' 送检论文和送审论文
			destFile2=fileid&"1."&fileExt2
			destPath2=uploadPath&"\"&destFile2
			file2.SaveAs destPath2

			If review_status=rsDetectPaperUploaded Then
				sqlDetect="EXEC spDeleteDetectResult "&rs("ID")&","&toSqlString(rs3("THESIS_FILE"))&";"
			End If
			sqlDetect=sqlDetect&"EXEC spAddDetectResult "&rs("ID")&","&toSqlString(destFile)&",NULL,NULL,NULL,1;"
			rs3("THESIS_FILE")=destFile
			rs3("THESIS_FILE2")=destFile2
			rs3("KEYWORDS")=new_keywords_ch
			rs3("KEYWORDS_EN")=new_keywords_en
			rs3("REVIEW_TYPE")=new_review_type
			rs3("REVIEW_STATUS")=rsDetectPaperUploaded
			If rs("DETECT_COUNT")=0 Then
				rs3("SUBMIT_REVIEW_TIME")=Null
				rs3("DETECT_APP_EVAL")=Null
				rs3("REVIEW_APP_EVAL")=Null
			End If
		Case sectionUploadDefence ' 答辩论文
			rs3("THESIS_FILE3")=destFile
			rs3("REVIEW_STATUS")=rsDefencePaperUploaded
			rs3("TUTOR_MODIFY_EVAL")=Null
		Case sectionUploadInstructReview ' 教指委盲评论文
			If review_status=rsInstructReviewPaperUploaded Then
				sqlDetect="EXEC spDeleteDetectResult "&rs("ID")&","&toSqlString(rs3("THESIS_FILE4"))&";"
			End If
			sqlDetect=sqlDetect&"EXEC spAddDetectResult "&rs("ID")&","&toSqlString(destFile)&",NULL,NULL,NULL,2;"
			rs3("THESIS_FILE4")=destFile
			rs3("REVIEW_STATUS")=rsInstructReviewPaperUploaded
		Case sectionUploadFinal ' 定稿论文
			rs3("THESIS_FILE5")=destFile
			rs3("REVIEW_STATUS")=rsFinalPaperUploaded
		End Select
		rs3.Update()
		CloseRs rs3

		If Len(sqlDetect) Then
			conn.Execute sqlDetect
		End If
		writeLog Format("学生[{0}]上传[{1}]。",Session("Stuname"),arrStuOprName(section_id))
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
<% useStylesheet "student" %>
<% useScript "jquery" %>
</head>
<body>
<center><br/><b>上传专业硕士论文</b><br/><br/><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=destFile%>" />
</form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%>
<script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></center></body></html><%
End Select
CloseRs rs
CloseRs rsStu
CloseConn conn
%>