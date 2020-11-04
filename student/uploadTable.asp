<!--#include file="../inc/global.inc"-->
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim is_new_dissertation:is_new_dissertation=False
Dim activity_id,section_id,time_flag,uploadable
Dim conn,rs,sql,count

activity_id=0
section_id=0
uploadable=False
stu_type=Session("StuType")

Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,2) AS STAT_TEXT FROM ViewDissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	section_id=sectionUploadKtbgb
	task_progress=tpNone
	str_keywords_ch="''"
	str_keywords_en="''"
Else
	paper_id=rs("ID")
	activity_id=rs("ActivityId")
	' 表格审核进度
	task_progress=rs("TASK_PROGRESS")
	Select Case task_progress
	Case tpNone,tpTbl1Uploaded,tpTbl1Unpassed	' 开题报告
		section_id=sectionUploadKtbgb
		is_generated=Not IsNull(rs("TABLE_FILE1"))
		filetype=1
	Case tpTbl1Passed,tpTbl2Uploaded,tpTbl2Unpassed	' 中期考核表
		section_id=sectionUploadZqkhb
		is_generated=Not IsNull(rs("TABLE_FILE2"))
		filetype=3
	Case tpTbl2Passed,tpTbl3Uploaded,tpTbl3Unpassed	' 预答辩意见书
		section_id=sectionUploadYdbyjs
		is_generated=Not IsNull(rs("TABLE_FILE3"))
		filetype=5
	Case tpTbl3Passed,tpTbl4Uploaded,tpTbl4Unpassed	' 答辩审批材料
		review_status=rs("REVIEW_STATUS")
		If review_status>=rsReviewEval Then
			section_id=sectionUploadSpclb
			is_generated=Not IsNull(rs("TABLE_FILE4"))
			filetype=7
		End If
	Case tpTbl4Passed ' 答辩审批材料审核通过
		section_id=sectionUploadSpclb
		is_generated=True
		filetype=7
	End Select
	subject_ch=rs("THESIS_SUBJECT")
	subject_en=rs("THESIS_SUBJECT_EN")
	sub_research_field=rs("RESEARCHWAY_NAME")
	str_keywords_ch="'"&Join(Split(toPlainString(rs("KEYWORDS")),"；"),"','")&"'"
	str_keywords_en="'"&Join(Split(toPlainString(rs("KEYWORDS_EN")),"；"),"','")&"'"
End If
If section_id<>0 Then
	is_new_dissertation=section_id=sectionUploadKtbgb Or stu_type=7 And section_id=sectionUploadYdbyjs
	If rs.EOF Then
		uploadable=True
	ElseIf Not isActivityOpen(rs("ActivityId")) Then
		time_flag=-3
	Else
		Set current_section=getSectionInfo(rs("ActivityId"), stu_type, section_id)
		time_flag=compareNowWithSectionTime(current_section)
		uploadable=time_flag=0
		If section_id<=sectionUploadYdbyjs Then
			is_tbl_thesis_uploaded=Not IsNull(rs("TBL_THESIS_FILE"&section_id))
		End If
	End If
End If
step=Request.QueryString("step")
Select Case step
Case vbNullstring ' 填写信息页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格文件</title>
<% useStylesheet "student", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "upload", "fillInTable", "keywordList" %>
</head>
<body>
<center><font size=4><b>上传表格文件</b></font>
<form id="fmDissertation" action="?step=1" method="post" enctype="multipart/form-data">
<table class="form" width="1000"><tr><td class="summary"><%
	If Not uploadable Then
		If time_flag=-3 Then
%><p><span class="tip">当前评阅活动【<%=rs("ActivityName")%>】已关闭，不能提交表格！</span></p><%
		ElseIf time_flag=-2 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节已关闭，不能提交表格！</span></p><%
		ElseIf time_flag<>0 Then
%><p><span class="tip">【<%=current_section("Name")%>】环节开放时间为<%=toDateTime(current_section("StartTime"),1)%>至<%=toDateTime(current_section("EndTime"),1)%>，当前不在开放时间内，不能提交表格！</span></p><%
		Else
%><p><span class="tip">当前状态为【<%=rs("STAT_TEXT")%>】，没有需提交的表格！</span></p><%
		End If
	Else
%><p>当前上传的是：<span style="color:#ff0000;font-weight:bold"><%=arrStuOprName(section_id)%></span></p><%
		If is_new_dissertation Then %>
<p>请选择您要参加的评阅活动：
<input id="activity_id" class="easyui-combobox" name="activity_id"
    data-options="valueField: 'id',
	textField: 'name',
	editable: false,
	prompt: '【请选择】',
	width: 300,
	panelHeight: 100,<%
	If activity_id<>0 Then %>
	value: <%=activity_id%>,<%
	End If %>
	url: '../api/get-attendable-activities',
	loadFilter: Common.curryLoadFilter(Array.prototype.reverse)"></p><%
		End If %>
<p>请选择要上传的文件，并点击&quot;提交&quot;按钮：</p><%
	End If %></td></tr>
<tr><td align="center">
<table class="form">
<tr><td><p>论文题目：《<input type="text" name="subject_ch" size="100" value="<%=subject_ch%>" />》</p>
<p>（英文）：&nbsp;<input type="text" name="subject_en" size="100" maxlength="200" value="<%=subject_en%>" /></p><%
	If is_new_dissertation Then %>
<p>工程领域：<select name="research_field_select"></select><input type="hidden" name="research_field" /></p>
<p>研究方向：<select name="sub_research_field_select"></select><input type="hidden" name="sub_research_field" />
<input type="text" name="custom_sub_research_field" size="20" placeholder="请输入…" value="<%=sub_research_field%>" /></p><%
	End If %>
<p>文件名：<input type="file" name="table_file" size="50" title="<%=arrStuOprName(section_id)%>" /><br/><span class="tip">Word&nbsp;或&nbsp;RAR&nbsp;格式，超过20M请先压缩成rar文件再上传，否则上传不成功</span></p></td></tr><%
	If is_new_dissertation Then %>
<tr><td><table width="420">
<tr><td colspan="3"><p>关键词（不少于三个，左边为中文，右边为英文）：<input type="hidden" name="keyword_ch_all" /><input type="hidden" name="keyword_en_all" /></td></tr>
<tr class="keywordpair"><td><p><input type="text" name="keyword_ch" class="keyword" size="10" placeholder="请输入……" /></p></td>
<td><p><input type="text" name="keyword_en" class="keyword" size="20" placeholder="Please input..." /></p></td>
<td><a class="linkRemove" href="#" title="删除此关键字"><img src="../images/student/remove.png" width="20" height="20" /></a></td></tr>
<tr><td colspan="3"><a class="linkAdd" href="#" title="增加关键字"><img src="../images/student/add.png" width="20" height="20" /></a></p></td></tr><%
	End If %>
</table></td></tr><%
	If section_id>0 And section_id<=sectionUploadYdbyjs And uploadable And Not is_tbl_thesis_uploaded Then %>
<tr><td align="center"><span class="tip">提示：您目前尚未上传<%=arrTblThesis(section_id)%>，<a href="uploadTablePaper.asp">点击这里上传。</a></span></td></tr><%
	End If %>
<tr><td align="center"><p><%
	If uploadable Then
%><input type="submit" id="btnsubmit" value="提 交" />&nbsp;<%
	End If
	If is_generated Then
%><input type="button" id="btndownload" value="下载打印" />&nbsp;<%
	End If
%><input type="button" name="btnreturn" value="返回首页" onclick="location.href='home.asp'" /></p></td></tr></table>
</td></tr></table></form></center>
<script type="text/javascript"><%
	If is_new_dissertation Then
		If str_keywords_ch<>"''" And str_keywords_en<>"''" Then %>
	$().ready(function() {
		setKeywords([<%=str_keywords_ch%>],[<%=str_keywords_en%>]);
	});<%
		End If
		If uploadable Then %>
	$('select[name="sub_research_field_select"]').change(function(){
		$('input[name="sub_research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
		var $custom_field=$('input[name="custom_sub_research_field"]');
		if(this.value=='-1')
			$custom_field.show().focus();
		else
			$custom_field.hide();
	});
	$('select[name="research_field_select"]').change(function(){
		initSubResearchFieldSelectBox($('select[name="sub_research_field_select"]'),$(this),this.value, '<%=sub_research_field%>');
		$('select[name="sub_research_field_select"]').change();
		$('input[name="research_field"]').val(!this.value.length?'':$(this).find('option:selected').text());
	});
	initResearchFieldSelectBox($('select[name="research_field_select"]'),<%=stu_type%>);<%
		End If
	End If %>
	$('form').submit(function(event) {<%
	If is_new_dissertation And uploadable Then %>
		if(!checkKeywords()) {
			event.preventDefault();
			return false;
		}<%
	End If %>
		var valid=checkIfWordRar(this.table_file);
		if(valid) submitUploadForm(this); else return false;
	});
	$('input[name="table_file"]').change(function(){if(this.value.length)checkIfWordRar(this);});<%
	If Not uploadable Then %>
	$('input[name="table_file"]').attr('readOnly',true);<%
	End If %>
	$(':button#btndownload').click(
		function() {
			window.location.href='fetchDocument.asp?tid=<%=paper_id%>&type=<%=filetype%>';
		}
	);
</script></body></html><%
Case 1	' 上传进程

	If time_flag=-3 Then
		bError=True
		errMsg=Format("当前评阅活动【{0}】已关闭，不能提交表格！", rs("ActivityName"))
	ElseIf time_flag=-2 Then
		bError=True
		errMsg=Format("【{0}】环节已关闭，不能上传表格文件！",current_section("Name"))
	ElseIf time_flag<>0 Then
		bError=True
		errMsg=Format("【{0}】环节开放时间为{1}至{2}，当前不在开放时间内，不能上传表格文件！",_
			current_section("Name"),_
			toDateTime(current_section("StartTime"),1),_
			toDateTime(current_section("EndTime"),1))
	ElseIf Not uploadable Then
		bError=True
		errMsg="当前状态为【"&rs("STAT_TEXT")&"】，没有需提交的表格！"
	End If
	If bError Then
		CloseRs rs
		CloseConn conn
		showErrorPage errMsg, "提示"
	End If

	Dim fso,Upload,table_file
	Dim new_subject_ch,new_subject_en
	Dim keyword_ch_all,keyword_en_all
	Dim research_field_select,sub_research_field_select,sub_research_field,custom_sub_research_field

	Set Upload=New ExtendedRequest
	activity_id=Upload.Form("activity_id")
	new_subject_ch=Upload.Form("subject_ch")
	new_subject_en=Upload.Form("subject_en")
	keyword_ch_all=Upload.Form("keyword_ch_all")
	keyword_en_all=Upload.Form("keyword_en_all")
	research_field_select=Upload.Form("research_field_select")
	sub_research_field_select=Upload.Form("sub_research_field_select")
	sub_research_field=Upload.Form("sub_research_field")
	custom_sub_research_field=Upload.Form("custom_sub_research_field")
	Set table_file=Upload.File("table_file")
	Set fso=CreateFSO()

	' 检查上传目录是否存在
	strUploadPath=Server.MapPath("upload")
	If Not fso.FolderExists(strUploadPath) Then fso.CreateFolder(strUploadPath)
	table_file_ext=LCase(table_file.FileExt)
	If is_new_dissertation Then
		If Len(activity_id)=0 Then
			bError=True
			errMsg="请选择要参加的评阅活动！"
		ElseIf stu_type = 5 And Len(research_field_select)=0 Then
			bError=True
			errMsg="请选择工程领域！"
		ElseIf Len(sub_research_field_select)=0 Then
			bError=True
			errMsg="请选择研究方向！"
		ElseIf UBound(Split(keyword_ch_all, ", "))<2 Then
			bError=True
			errMsg="请输入至少三个论文关键字（中文）！"
		ElseIf UBound(Split(keyword_en_all, ", "))<2 Then
			bError=True
			errMsg="请输入至少三个论文关键字（英文）！"
		ElseIf sub_research_field_select="-1" And Len(custom_sub_research_field)=0 Then
			custom_sub_research_field="其他"
		End If
	ElseIf InStr("doc docx rar",table_file_ext)=0 Then	' 不被允许的文件类型
		bError=True
		errMsg="所上传的不是 Word 文件或 RAR 压缩文件！"
	ElseIf Len(new_subject_ch)=0 Then
		bError=True
		errMsg="请填写论文题目！"
	ElseIf Len(new_subject_en)=0 Then
		bError=True
		errMsg="请填写论文题目（英文）！"
'	ElseIf file.FileSize>10485760 Then
'		filesize=Round(file.FileSize/1048576,2)
'		bError=True
'		errMsg="文件大小为 "&filesize&"MB，已超出限制(10MB)！"
	End If
	If Not bError Then
		byteFileSize=0
		' 生成日期格式文件名
		fileid=timestamp()
		strDestTableFile=fileid&"."&table_file_ext
		destPath=strUploadPath&"\"&strDestTableFile
		byteFileSize=table_file.FileSize
		' 保存表格文件
		table_file.SaveAs destPath
	End If
	Set fso=Nothing
	Set table_file=Nothing
	Set Upload=Nothing

	If Not bError Then
		Dim arrTableFieldName,arrNewTaskProgress
		arrTableFieldName=Array("","TABLE_FILE1","TABLE_FILE2","TABLE_FILE3","TABLE_FILE4")
		arrNewTaskProgress=Array(0,tpTbl1Uploaded,tpTbl2Uploaded,tpTbl3Uploaded,tpTbl4Uploaded)
		' 关联到数据库
		sql="SELECT * FROM Dissertations WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
		GetRecordSet conn,rs3,sql,count
		If rs3.EOF Then
			' 添加记录
			rs3.AddNew()
		End If
		If is_new_dissertation Then	' 新论文记录，录入论文基本信息
			rs3("STU_ID")=Session("Stuid")
			rs3("ActivityId")=activity_id
			rs3("REVIEW_STATUS")=rsNone
			rs3("REVIEW_RESULT")="5,5,6"
			rs3("REVIEW_LEVEL")="0,0"
			rs3("KEYWORDS")=Replace(keyword_ch_all,", ","；")
			rs3("KEYWORDS_EN")=Replace(keyword_en_all,", ","；")
			If sub_research_field_select="-1" Then
				rs3("RESEARCHWAY_NAME")=custom_sub_research_field
			ElseIf Len(sub_research_field) Then
				rs3("RESEARCHWAY_NAME")=sub_research_field
			End If
		End If
		rs3("THESIS_SUBJECT")=new_subject_ch
		rs3("THESIS_SUBJECT_EN")=new_subject_en
		rs3(arrTableFieldName(section_id))=strDestTableFile
		rs3("TASK_PROGRESS")=arrNewTaskProgress(section_id)
		rs3.Update()
		CloseRs rs3

		writeLog Format("学生[{0}]上传[{1}]。",Session("Stuname"),arrStuOprName(section_id))
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>上传表格文件</title>
<% useStylesheet "student" %>
<% useScript "jquery" %>
</head>
<body><%
	If Not bError Then %>
<form id="fmFinish" action="home.asp" method="post">
<input type="hidden" name="filename" value="<%=strDestTableFile%>" />
</form>
<script type="text/javascript">alert("上传成功！");$('#fmFinish').submit();</script><%
	Else
%><script type="text/javascript">alert("<%=errMsg%>");history.go(-1);</script><%
	End If
%></body></html><%
End Select
CloseRs rs
CloseConn conn
%>