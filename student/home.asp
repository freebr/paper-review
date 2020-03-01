<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
Dim bTableFilledIn:bTableFilledIn=Array(0,False,False,False,False)
Dim bTblThesisUploaded:bTblThesisUploaded=Array(0,False,False,False)
Dim arrTableStat(4)
Dim review_result(2),bReviewFileVisible(1)
Dim defence_member,defence_members,defence_memo,defence_eval
Dim defence_result,grant_degree_result

sem_info=getCurrentSemester()
task_progress=0
stu_type=Session("StuType")

Connect conn
sql="SELECT * FROM ViewDissertations_student WHERE STU_ID="&Session("Stuid")&" ORDER BY ActivityId DESC"
GetRecordSetNoLock conn,rs,sql,count
If rs.EOF Then
	task_prog_text="未上传开题报告"
	review_stat_text="未上传送检论文"
Else
	thesisID=rs("ID")
	task_progress=rs("TASK_PROGRESS")
	thesis_files=Array(rs("THESIS_FILE").Value,rs("THESIS_FILE2").Value,rs("THESIS_FILE3").Value,rs("THESIS_FILE4").Value)
	reproduct_ratio=toNumericString(rs("REPRODUCTION_RATIO"))
	instruct_review_reproduct_ratio=toNumericString(rs("INSTRUCT_REVIEW_REPRODUCTION_RATIO"))
	detect_count=rs("DETECT_COUNT")
	bReviewFileVisible(0)=(rs("ReviewFileDisplayStatus1") And 2)<>0
	bReviewFileVisible(1)=(rs("ReviewFileDisplayStatus2") And 2)<>0
	bAllReviewFileVisible=bReviewFileVisible(0) And bReviewFileVisible(1)
	defence_member=rs("DEFENCE_MEMBER")
	defence_result=rs("DEFENCE_RESULT")
	defence_eval=rs("DEFENCE_EVAL")
	grant_degree_result=rs("GRANT_DEGREE_RESULT")
	For i=1 To UBound(arrTableStat)
		j=(i-1)*3+1
		k=i*3
		If task_progress<j Then
			arrTableStat(i)=0
		ElseIf task_progress>k Then
			arrTableStat(i)=3
		Else
			arrTableStat(i)=(task_progress-1)Mod 3+1
			If arrTableStat(i)=2 Then
				bTableNotPassed=True
			End If
		End If
	Next
	For i=1 To 4
		bTableFilledIn(i)=Not IsNull(rs("TABLE_FILE"&i))
		If i<4 Then bTblThesisUploaded(i)=Not IsNull(rs("TBL_THESIS_FILE"&i))
	Next
	review_status=rs("REVIEW_STATUS")
	review_type=rs("REVIEW_TYPE")
	task_prog_text=rs("TaskProgressText")
	review_stat_text=rs("ReviewStatusText")
	If Not IsNull(rs("REVIEW_RESULT")) Then
		arrRevRet=Split(rs("REVIEW_RESULT"),",")
		For i=0 To UBound(arrRevRet)
			review_result(i)=Int(arrRevRet(i))
		Next
	End If
End If
Function showStepInfo(stepDisplay,stepCounter,is_hidden)
	If stepDisplay=rsRedetectPassed And detect_count>1 And (review_status=rsAgreedReview Or review_status=rsRefusedReview) Then
		showStepInfo=True
	ElseIf review_status<>stepDisplay And _
		(stepDisplay=rsRefusedDetect Or _
		stepDisplay=rsDetectUnpassed Or _
		stepDisplay=rsRedetectUnpassed Or _
		stepDisplay=rsRedetectPassed Or _
		stepDisplay=rsRefusedReview Or _
		stepDisplay=rsRefusedDefence Or _
		stepDisplay=rsRefusedInstructReview) Then
		showStepInfo=False
		Exit Function
	Else
		showStepInfo=True
	End If
	Dim i,className,className_seqline
	Dim audit_info
	If is_hidden Then
		className="hidden"
	ElseIf review_status=stepDisplay Then
		className="current"
	End If
	If Len(className) Then className="class="""&className&""""
	If stepDisplay>1 Then className_seqline="class=""seqline""" Else className_seqline=""
%><tr <%=className%> id="step<%=stepDisplay%>"><td class="stepicon"></td><td class="steptext"><p class="stepname"><%=arrStep(stepDisplay)%></p></td></tr>
<tr <%=className%>><td <%=className_seqline%>></td><td class="steptext"><p class="stepcontent"><%
	Select Case stepDisplay
	Case rsDetectPaperUploaded
%><span class="section-detail">在导师同意检测前，您可以重复上传送检论文文件；导师仅能看到您最新上传的论文。</span><%
	Case rsRefusedDetect
		audit_info=getAuditInfo(thesisID,thesis_files(0),auditTypeDetectReview)
%><span class="section-detail">导师不同意您的论文进行查重，请修改论文后重新上传。
<br/>送检意见：<%=toPlainString(audit_info(0)("Comment"))%></span><%
	Case rsAgreedDetect
%><span class="section-detail">导师已同意您的论文进行查重。</span><%
	Case rsDetectUnpassed,rsRedetectUnpassed,rsRedetectPassed
%><span class="section-detail"><%
		Dim filetype
		Select Case stepDisplay
		Case rsDetectUnpassed
			filetype=14
%>经过检测，您的送检论文文字复制比为&nbsp;<%=reproduct_ratio%>%，不符合学院送检论文重复率低于10%的要求，请对论文修改后重新上传进行二次检测。<%
		Case rsRedetectUnpassed
			filetype=15
%>经过二次检测，您的送检论文文字复制比为&nbsp;<%=reproduct_ratio%>%，不符合学院送检论文重复率低于10%的要求。<%
		Case Else
			filetype=15
%>您的论文已通过二次查重检测，请等候导师同意送审。<br/>检测结果摘要：经图书馆检测，学位论文文字复制比为&nbsp;<%=reproduct_ratio%>%。<%
		End Select
%><br/><a class="resc" href="fetchDocument.asp?tid=<%=thesisID%>&type=<%=filetype%>" target="_blank">点此下载检测报告</a></span><%
	Case rsRefusedReview
		audit_info=getAuditInfo(thesisID,thesis_files(1),auditTypeReviewApp)
%><span class="section-detail">导师不同意您的论文送审，请对照导师意见修改送审论文后重新上传。<br/>送审意见：<%=toPlainString(audit_info(0)("Comment"))%></span><%
	Case rsAgreedReview
%><span class="section-detail"><%
		audit_info=getAuditInfo(thesisID,thesis_files(1),auditTypeReviewApp)
		If detect_count>1 Then
			filetype=15
%>导师已于&nbsp;<%=toDateTime(rs("SUBMIT_REVIEW_TIME"),1)&" "&toDateTime(rs("SUBMIT_REVIEW_TIME"),4)%>&nbsp;同意您的论文送审申请，教务员将匹配专家对您的论文进行评阅。<%
		Else
			filetype=14
%>您的论文已通过查重检测，教务员将匹配专家对您的论文进行评阅。<br/>检测结果摘要：经图书馆检测，学位论文文字复制比为&nbsp;<%=reproduct_ratio%>%。<%
		End If
%><br/>送审意见：<%=toPlainString(audit_info(0)("Comment"))%>
<br/><a class="resc" href="fetchDocument.asp?tid=<%=thesisID%>&type=<%=filetype%>" target="_blank">点此下载检测报告</a></span><%
	Case rsMatchedReviewer
%><span class="section-detail">教务员已为您的论文匹配了评阅专家，正在对您的论文进行评阅，请耐心等候评阅结果。</span><%
	Case rsReviewed
%><span class="section-detail">专家已完成论文评阅，请等候导师进行确认。</span><%
	Case rsReviewEval
%><span class="section-detail">专家已完成论文评阅，请按照评阅书意见对论文进行修改，然后上传答辩论文。</span><br/><%
		If bReviewFileVisible(0) Then
%>评阅意见&nbsp;1：【<%=getReviewResultText(review_result(0))%>】&nbsp;<%
		End If
		If bReviewFileVisible(1) Then
%>评阅意见&nbsp;2：【<%=getReviewResultText(review_result(1))%>】&nbsp;<%
		End If
		If bAllReviewFileVisible Then
%>总体评价：【<%=getFinalResultText(review_result(2))%>】<br/><%
		End If
		For i=0 To 1
			If arrRevRet(i)<>5 And bReviewFileVisible(i) Then	' 该专家已评阅且评阅书开放显示
				If i=1 Then Response.Write "&emsp;"
%><a class="resc" href="fetchDocument.asp?tid=<%=thesisID%>&type=<%=17+i%>" target="_blank">点击下载第<%=i+1%>份评阅书</a></span><%
			End If
		Next
	Case rsDefencePaperUploaded
%><span class="section-detail">您已上传答辩论文，请等候导师审核。<%
		If task_progress<tpTbl4Uploaded Then
%>待导师审核后，将同意答辩的意见嵌入《学位论文答辩及授予学位审批材料》中打印并找导师签字。<%
		End If
%></span><%
	Case rsRefusedDefence
		audit_info=getAuditInfo(thesisID,thesis_files(2),auditTypeDefence)
%><span class="section-detail">您的答辩论文未获导师<%=rs("TUTOR_NAME")%>审核通过，请修改论文后重新上传。审核意见如下：<br/>&emsp;&emsp;<%=toPlainString(audit_info(0)("Comment"))%></span><%
	Case rsAgreedDefence
		audit_info=getAuditInfo(thesisID,thesis_files(2),auditTypeDefence)
%><span class="section-detail">导师<%=rs("TUTOR_NAME")%>已审核通过您的答辩论文。审核意见如下：<br/>&emsp;&emsp;<%=toPlainString(audit_info(0)("Comment"))%></span><%
	Case rsDefenceEval
%><span class="section-detail">答辩委员会已对您的论文提出了如下修改意见，请根据意见修改并上传教指委盲评论文。<br/><%=toPlainString(defence_eval)%></span><%
	Case rsInstructReviewPaperUploaded
%><span class="section-detail">您已上传教指委盲评论文，请等候导师审核。</span><%
	Case rsRefusedInstructReview
		audit_info=getAuditInfo(thesisID,thesis_files(3),auditTypeInstructReviewDetect)
%><span class="section-detail">导师不同意您的教指委盲评论文进行查重，请修改论文后重新上传。<br/>审核意见：<%=toPlainString(audit_info(0)("Comment"))%></span><%
	Case rsAgreedInstructReview
%><span class="section-detail">导师已同意您的教指委盲评论文进行查重。</span><%
	Case rsInstructReviewPaperDetected
%><span class="section-detail">您的教指委盲评论文已完成查重，请等候教务员为您的论文匹配教指委委员。<br/>检测结果摘要：经图书馆检测，学位论文文字复制比为&nbsp;<%=instruct_review_reproduct_ratio%>%。
<br/><a class="resc" href="fetchDocument.asp?tid=<%=thesisID%>&type=16" target="_blank">点此下载检测报告</a></span><%
	Case rsMatchedInstructMember
%><span class="section-detail">教务员已为您的论文匹配教指委委员，请等候委员审核。</span><%
	Case rsInstructEval
		audit_info=getAuditInfo(thesisID,thesis_files(3),auditTypeInstructReview)
%><span class="section-detail">教指委已对您的论文提出了如下修改意见：
<br/>第一位委员的意见：<br/><%=toPlainString(audit_info(0)("Comment"))%>
<br/>第二位委员的意见：<br/><%=toPlainString(audit_info(1)("Comment"))%><%
		If Not IsNull(rs("DEGREE_MODIFY_EVAL")) Then %>
<br/>学院学位评定分会修改意见：<br/><%=toPlainString(rs("DEGREE_MODIFY_EVAL"))%></span><%
		End If
	Case rsFinalPaperUploaded
%><span class="section-detail">您已上传定稿论文。<%
	End Select
%></p></td></tr><%
End Function
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>系统首页</title>
<% useStylesheet "student" %>
<% useScript "jquery", "common", "paper" %>
<style type="text/css">
	td.modtitle { height:20;border:1px solid gainsboro }
	td.modcontent { padding-left:20px;padding-top:10px;background:url(../images/student/modback.png) repeat }
	p.stepname { margin-top:10px;margin-bottom:0 }
	p.stepcontent { margin:0 0 5px 0;padding-left:10px }
	p.tutoreval { padding-left:20px;text-align:left;color:#0000ff;font-size:10pt }
	p.defenceresult { padding:20px 0 }
	p.defenceresult span { color:#ff0000;font-weight:bold }
	table.taskprogress { margin:10px 0 }
	table.taskprogress td { display:inline-block;table-layout:fixed;padding:10px;border:1px solid #3399dd;
													background-color:#fffddf;font-size:10pt;text-align:center }
	table.taskprogress td.step { height: 100px }
	table.taskprogress span.s0 { color:#cccccc }
	table.taskprogress span.s1 { color:#0000cc }
	table.taskprogress span.s2 { color:#cc0000 }
	table.taskprogress span.s3 { color:#00cc00 }
	table.taskprogress p.smalltext { text-align:left;color:#00aa00;font-size:9pt;font-weight:bold }
	table.seqchart { width:100%; border-spacing:0; table-layout:fixed }
	table.seqchart td.stepicon { padding:1px 0; width:32px; background-clip: border-box }
<%
	For i=1 To review_status %>
	table.seqchart tr#step<%=i%> td.stepicon { background:url(../images/student/step<%=i%>.png) no-repeat }<%
	Next %>
	table.seqchart td.steptext { padding:1px 0; width:900px; height:100% }
	table.seqchart td.seqline { background:url(../images/student/seqline.png) repeat-y 13px 0; height:20px; min-height:20px }
	table.seqchart td.seqmore { background:url(../images/student/seqmore.png) no-repeat left center; height:20px; min-height:20px; cursor:pointer }
	table.seqchart tr.hidden { visibility:hidden }
	table.seqchart tr.half-hidden {
 		-webkit-text-fill-color: transparent;
 	  -webkit-background-clip: text;
 		background-image: linear-gradient(rgba(0,0,0,0.8),rgba(0,0,0,0));
  	filter: alpha(opacity=1,finishopacity=0,style=1,startx=0,starty=0,finishx=0,finishy=10);
  	cursor: pointer;
	}
	table.seqchart tr.half-hidden td.stepicon {
		mix-blend-mode: darken;
	}
	table.seqchart tr.current { background-color:#ecffec }
	table.seqchart tr.current p.stepname { color:#0000cc;font-weight:bold }
	table.seqchart tr.current { background-color:#ecffec }
	div#defenceplan { margin:0 20px 10px 0;width:1000px;border:1px solid #3399dd;
										background-color:#fffddf }
	div#defenceplan table { table-layout:fixed;width:100%;text-align:center }
	div#defenceplan table thead tr { background-color:#3399dd;color:white }
	div#defenceplan table td { padding:5px }
	ul.filelist { padding:0 }
	ul.filelist li { display:inline-block;width:70px;padding:5px 0px;list-style:none;text-align:center;vertical-align:top }
	span.filedesc { color:#666666 }
	a.fileitem { display:inline-block;padding:3px;text-align:center }
	a.fileitem:visited { background-color:none }
	a.fileitem:link { background-color:none }
	a.fileitem:hover { background-color:#BFF5FF;color:0;text-decoration:none }
</style>
</head>
<body>
<center><font size=4><b>欢迎使用工商管理学院专业学位论文评阅系统</b></font>
<%
	If stu_type=6 Then
		If review_status<>rsNone And (review_type=0 Or review_type=1) Then
			sql="SELECT * FROM ReviewTypes WHERE LEN(THESIS_FORM)>0 AND TEACHTYPE_ID="&stu_type
			GetRecordSetNoLock conn,rs2,sql,count
%><form method="post" action="setPaperForm.asp">
<input type="hidden" name="tid" value="<%=thesisID%>" />
<p><span class="tip">您还没有选择所撰写的论文形式，请在此选择并提交：</span>
<select id="thesisform" name="thesis_form" style="width:350px"><option value="0">请选择……</option><%
			Do While Not rs2.EOF
%><option value="<%=rs2("ID")%>"><%=rs2("THESIS_FORM")%></option><%
				rs2.MoveNext()
			Loop
			CloseRs rs2
%></select>
<input type="submit" name="btnsubmit" value="提 交" /></p></form>
<%
		End If
	End If
%>
<table width="100%" cellpadding="0" cellspacing="1" style="font-size:10pt;line-height:20px">
<tr bgcolor="#E4E8EF"><td class="modtitle"><img src="../images/student/bullet.gif">&nbsp;<b>当前评阅进度&nbsp;—【<%=review_stat_text%>】</b></td></tr>
<tr><td class="modcontent" width="100%" valign="top">
<table class="seqchart"><%
		Dim is_hidden,maxVisibleSteps:maxVisibleSteps=3
		j=1
		' 显示最新步骤
		For i=review_status To 1 Step -1
			is_hidden=review_status-i>=maxVisibleSteps
			If showStepInfo(i,j,is_hidden) Then j=j+1
		Next
		If is_hidden Then
%><tr><td class="seqmore" colspan="2" title="点击显示早前已完成的评阅流程"></td></tr><%
		End If %>
</table><%
		If Not rs.EOF Then
			' 显示答辩安排
			If Not IsNull(defence_member) Then
				defence_members=Split(defence_member,"|")
				defence_memo=rs("MEMO")
				If IsNull(defence_memo) Then defence_memo="-"
%><hr/><p>您的论文答辩安排如下：
<div id="defenceplan"><table cellspacing="0" cellpadding="1">
<thead><tr style="font-weight:bold"><td width="120"><p>答辩时间</p></td><td width="100"><p>答辩地点</p></td>
<td width="70"><p>答辩主席</p></td><td width="150"><p>答辩委员</p></td><td width="70"><p>答辩秘书</p></td><td><p>答辩委员工作单位</p></td></tr></thead>
<tbody><tr><td><p><%=rs("DEFENCE_TIME")%></p></td><td><p><%=rs("DEFENCE_PLACE")%></p></td>
<td><p><%=defence_members(0)%></p></td><td><p><%=defence_members(1)%></p></td><td><p><%=defence_members(2)%></p></td>
<td><p><%=toPlainString(defence_memo)%></p></td></tbody></table></div></p><%
			End If
			
			' 显示答辩成绩
			If Not IsNull(defence_result) And defence_result<>0 Then %>
	<hr/><p class="defenceresult"><span>您的答辩成绩为：<%=arrDefenceResult(defence_result)%>，答辩表决结果：<%=arrGrantDegreeResult(grant_degree_result)%></span></p><%
			End If
		End If
%></td></tr>
<tr bgcolor="#E4E8EF"><td class="modtitle"><img src="../images/student/bullet.gif">&nbsp;<b>相关表格审核进度&nbsp;—【<%=task_prog_text%>】</b></td></tr>
<tr><td class="modcontent" width="100%" valign="top">
<table class="taskprogress" width="600" cellpadding="0" cellspacing="0"><tr><%
		Dim arrFillInText:arrFillInText=Array("未填写","已填写")
		Dim arrUploadText:arrUploadText=Array("未上传","已上传")
		For i=1 To UBound(arrTable)
%><td class="step"><p><%=arrTable(i)%><br/><span class="s<%=arrTableStat(i)%>"><%=arrTableStatText(arrTableStat(i))%></span></p><%
%><p class="smalltext">表格：<%=arrFillInText(Abs(Int(bTableFilledIn(i))))%><%
			If i<=3 Then
				' 若最新环节未上传附加论文，则显示上传按钮，否则显示状态
				If Not bTblThesisUploaded(i) And (arrTableStat(i)=1 Or arrTableStat(i)=2) Then %>
<br/><input type="button" name="btnUploadTableThesis" value="上传<%=arrTblThesis(i)%>..." onclick="location.href='uploadTablePaper.asp'" /><%
				Else %>
<br/><%=arrTblThesis(i)%>：<%=arrUploadText(Abs(Int(bTblThesisUploaded(i))))%><%
				End If
			End If
%></p></td><%
		Next %></tr><%
		If bTableNotPassed Then
%><tr><td colspan="4" width="90%"><p align="left">【导师&nbsp;<%=rs("TUTOR_NAME")%>&nbsp;的审核意见】</p>
	<p class="tutoreval"><%=toPlainString(rs("TASK_EVAL"))%></p></td></tr><%
		End If %></table></td></tr>
<tr bgcolor="#E4E8EF"><td class="modtitle"><img src="../images/student/bullet.gif">&nbsp;<b>论文及相关文件</b></td></tr>
<tr><td class="modcontent" width="100%" height="180" valign="top"><%
	If rs.EOF Then %>
<p style="font-size:10pt">当前还没有上传或生成过任何文件！</b></p>
<%
	Else %>
<p style="font-size:10pt">评阅活动：<b><%=rs("ActivityName")%></b></p>
<p style="font-size:10pt">论文题目：<b><%=rs("THESIS_SUBJECT")%></b></p><p><ul class="filelist"><%
		Dim fso,file
		Set fso=Server.CreateObject("Scripting.FileSystemObject")
		For i=1 To UBound(arrFileListName)
			filename=rs(arrFileListField(i))
			If Not IsNull(filename) Then
				If i=17 Or i=18 Then
					' 根据评阅书显示设置决定是否显示文件
					If Not bReviewFileVisible(i-17) Then filename=""
				End If
				filepath=arrFileListPath(i)&"/"&filename
				fullfilepath=Server.MapPath(arrFileListPath(i))&"\"&filename
				If fso.FileExists(fullfilepath) Then
					Set file=fso.GetFile(fullfilepath)
					file_ext=fso.GetExtensionName(filename) %>
<li><a class="fileitem" href="fetchDocument.asp?tid=<%=thesisID%>&type=<%=i%>" target="_blank" title="大小：<%=toDataSizeString(file.Size)%>&#10;创建时间：<%=FormatDateTime(file.DateCreated,2)&" "&FormatDateTime(file.DateCreated,4)%>&#10;点击下载此文件"><img src="../images/student/<%=file_ext%>.png" title="<%=UCase(file_ext)%>格式" /><div><%=arrFileListName(i)%></div></a></li><%
					Set file=Nothing
				End If
			End If
		Next
		Set fso=Nothing
	End If
%></ul></p></td></tr></table></center>
<script type="text/javascript">
	$(document).ready(function() {
		$('table.seqchart tr.hidden').each(function(){$(this).attr({'finalHeight':$(this).height(),'height':1}).hide().css('visibility','visible');})
																 .eq(0).addClass('half-hidden').show();
		$('td.seqmore').click(function() {
			$('table.seqchart tr.hidden').show().each(function(){$(this).removeClass('hidden').removeClass('half-hidden').animate({'height':$(this).attr('finalHeight')},1000)});
			$(this).hide();
		});
		$('table.seqchart tr.hidden:eq(0)').click(function() {
			$('td.seqmore').click();
		});
	});
</script></body></html><%
CloseRs rs
CloseConn conn
%>