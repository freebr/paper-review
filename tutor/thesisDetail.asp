<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
If Not hasPrivilege(Session("Treadprivileges"),"I11") Then Response.Redirect("../error.asp?no-privilege")
curstep=Request.QueryString("step")
thesisID=Request.QueryString("tid")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
tid=Session("Tid")
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">参数无效。</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim table_file(4)
Connect conn
sql="SELECT *,dbo.getThesisStatusText(1,TASK_PROGRESS,3) AS STAT_TEXT1,dbo.getThesisStatusText(2,REVIEW_STATUS,3) AS STAT_TEXT2 FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
GetRecordSet conn,rs,sql,result
If result=0 Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If

Dim arrCaptionAgree:arrCaptionAgree=Array("","审核通过开题报告表，同意参加开题报告","审核通过中期检查表","审核通过预答辩申请表，同意参加预答辩","审核通过审批材料，同意参加答辩")
Dim arrCaptionNotAgree:arrCaptionNotAgree=Array("","审核不通过，不同意参加开题报告（延期3-6个月重新做开题报告）","审核不通过中期检查表","审核不通过预答辩申请表，不同意参加预答辩","审核不通过审批材料，不同意参加答辩")
Dim review_status,numReviewed,review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)

sql="SELECT * FROM VIEW_DETECT_RESULT_INFO WHERE THESIS_ID="&thesisID
GetRecordSet conn,rsDetect,sql,result
If rs("REVIEWER1")=tid Then
	reviewer=0
ElseIf rs("REVIEWER2")=tid Then
	reviewer=1
End If
If rs("TEACHTYPE_ID")=5 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
tutor_id=rs("TUTOR_ID")
review_app=rs("REVIEW_APP")
task_progress=rs("TASK_PROGRESS")
review_status=rs("REVIEW_STATUS")
stat_text1=rs("STAT_TEXT1")
stat_text2=rs("STAT_TEXT2")
bIsReviewVisible=(rs("REVIEW_FILE_STATUS") And 1)<>0
reproduct_ratio=rs("REPRODUCTION_RATIO")
defence_result=rs("DEFENCE_RESULT")
grant_degree=rs("GRANT_DEGREE")
opr=0
Select Case task_progress
Case tpNone
Case tpTbl1Uploaded:opr=1
Case tpTbl2Uploaded:opr=2
Case tpTbl3Uploaded:opr=3
Case tpTbl4Uploaded:opr=4
End Select
Select Case review_status
Case rsDetectThesisUploaded:opr=5
Case rsRedetectPassed:opr=6
Case rsMatchExpert:opr=9
Case rsReviewed:opr=7
Case rsModifyThesisUploaded:opr=8
End Select
If review_status=0 Then
	stat=stat_text1
ElseIf task_progress>=tpTbl4Uploaded Then
	stat=stat_text1&"，"&stat_text2
Else
	stat=stat_text2
End If
For i=1 To 4
	table_file(i)=rs("TABLE_FILE"&i)
Next
If Not IsNull(rs("THESIS_FILE")) Then
	thesis_file=rs("THESIS_FILE")
End If
If Not IsNull(rs("THESIS_FILE2")) Then
	thesis_file_review=rs("THESIS_FILE2")
End If
If Not IsNull(rs("THESIS_FILE3")) Then
	thesis_file_modified=rs("THESIS_FILE3")
End If
If Not IsNull(rs("THESIS_FILE4")) Then
	thesis_file_final=rs("THESIS_FILE4")
End If
If Not IsNull(rs("REVIEW_RESULT")) Then
	arr=Split(rs("REVIEW_RESULT"),",")
	For i=0 To UBound(arr)
		review_result(i)=Int(arr(i))
	Next
End If
If Not IsNull(rs("REVIEWER_EVAL_TIME")) Then
	arr2=Split(rs("REVIEWER_MASTER_LEVEL"),",")
	arr3=Split(rs("REVIEW_FILE"),",")
	arr4=Split(rs("REVIEWER_EVAL_TIME"),",")
	arr5=Split(rs("REVIEW_LEVEL"),",")
	For i=0 To 1
		reviewer_master_level(i)=Int(arr2(i))
		review_file(i)=arr3(i)
		review_time(i)=arr4(i)
		review_level(i)=Int(arr5(i))
	Next
	numReviewed=UBound(arr2)+1
End If
Select Case curstep
Case vbNullString	' 论文详情页面

	updateActiveTime Session("Tid")
	
	Dim tutor_modify_eval_title
	arrActionUrlList=Array("?tid="&thesisID&"&step=2","updateThesis.asp?tid="&thesisID,"../expert/thesisDetail.asp?tid="&thesisID&"&step=2")
	Select Case opr
	Case 1,2,3
		actionUrl1=arrActionUrlList(0)
		actionUrl2=arrActionUrlList(1)
	Case 4,5,6,8
		actionUrl1=arrActionUrlList(0)
		actionUrl2=actionUrl1
	Case 7
		actionUrl1=arrActionUrlList(1)
		actionUrl2=actionUrl1
	Case 9
		actionUrl1=arrActionUrlList(2)
		actionUrl2=vbNullString
	End Select
	If review_status>=rsModifyPassed Then
		tutor_modify_eval_title="导师同意答辩意见"
	ElseIf review_status=rsModifyUnpassed Then
		tutor_modify_eval_title="导师不同意答辩意见"
	Else
		tutor_modify_eval_title="导师对答辩论文的意见"
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/tutor.css" rel="stylesheet" type="text/css" />
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>查看论文信息</title>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>专业硕士论文详情<br/>论文当前状态：【<%=stat%>】</b></font><%
	If opr<>0 And nSystemStatus<>2 Then
%><p><span class="tip">温馨提示：评阅系统的开放时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，当前不在开放时间内，不能对论文执行操作。</span></p><%
	End If %>
<form id="fmDetail" action="<%=actionUrl%>" method="post">
<table class="tblform" width="800" cellspacing="1" cellpadding="3">
<tr><td>论文题目：&emsp;&emsp;&emsp;<input type="text" class="txt" name="subject" size="95%" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr>
<tr><td>（英文）：&emsp;&emsp;&emsp;<input type="text" class="txt" name="subject_en" size="85%" value="<%=rs("THESIS_SUBJECT_EN")%>" readonly /></td></tr>
<tr><td>作者姓名：&emsp;&emsp;&emsp;<input type="text" class="txt" name="author" size="18" value="<%=rs("STU_NAME")%>" readonly />&nbsp;
学号：<input type="text" class="txt" name="stuno" size="15" value="<%=rs("STU_NO")%>" readonly />&nbsp;
学位类别：<input type="text" class="txt" name="degreename" size="10" value="<%=rs("TEACHTYPE_NAME")%>" readonly />&nbsp;
学期：<input type="text" class="txt" name="periodid" size="6" value="<%=rs("PERIOD_ID")%>" readonly /></td></tr>
<tr><td>指导教师：&emsp;&emsp;&emsp;<input type="text" class="txt" name="tutorname" size="95%" value="<%=rs("TUTOR_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
<tr><td>领域名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
<tr><td>研究方向：&emsp;&emsp;&emsp;<input type="text" class="txt" name="researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td></tr>
<tr><td>论文关键词：&emsp;&emsp;<input type="text" class="txt" name="keywords_ch" size=85%" value="<%=rs("KEYWORDS")%>" readonly /></td></tr>
<tr><td>（英文）：&emsp;&emsp;&emsp;<input type="text" class="txt" name="keywords_en" size="85%" value="<%=rs("KEYWORDS_EN")%>" readonly /></td></tr>
<tr><td>院系名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="faculty" size="30%" value="工商管理学院" readonly />&nbsp;
班级：<input type="text" class="txt" name="class" size="51%" value="<%=rs("CLASS_NAME")%>" readonly /></td></tr><%
	If Not IsNull(rs("THESIS_FORM")) And Len(rs("THESIS_FORM")) Then %>
<tr><td>论文形式：&emsp;&emsp;&emsp;<input type="text" class="txt" name="thesisform" size="95%" value="<%=rs("THESIS_FORM")%>" readonly /></td></tr><%
	End If
	If review_status>=rsDetectUnpassed Then %>
<tr><td>学位论文文字复制比：<input type="text" class="txt" name="reproduct_ratio" size="10px" value="<%=reproduct_ratio%>%" readonly /></td></tr><%
	End If
	If task_progress>=tpTbl1Uploaded Then
		If Len(table_file(1)) Then %>
<tr><td>开题报告表：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=1" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
<tr><td>开题论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=2" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If task_progress>=tpTbl2Uploaded Then
		If Len(table_file(2)) Then %>
<tr><td>中期检查表：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=3" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
<tr><td>中期论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=4" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If task_progress>=tpTbl3Uploaded Then
		If Len(table_file(3)) Then %>
<tr><td>预答辩申请表：&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=5" target="_blank">点击下载</a></td></tr><%
		End If
		If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
<tr><td>预答辩论文：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=6" target="_blank">点击下载</a></td></tr><%
		End If
	End If
	If review_status>=rsDetectThesisUploaded And Len(thesis_file) Then %>
<tr><td>送检论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=8" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsDetectThesisUploaded Then %>
<tr><td>送检论文检测报告：<%
		If IsNull(rs("DETECT_REPORT")) Then %>
未上传<%
		Else %>
<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=12" target="_blank">点击下载</a><%
		End If
%></td></tr>
<tr><td>论文检测记录（按检测先后顺序）：<%
		If rsDetect.EOF Then
%>无<%
		Else
%><ul><%
			Dim index:index=1
			Dim hash,detect_time,detect_result
			Do While Not rsDetect.EOF
				hash=rsDetect("HASH").Value
				detect_time=rsDetect("DETECT_TIME").Value
				If IsNull(detect_time) Then detect_time="无"
				detect_result=rsDetect("RESULT").Value
				If IsNull(detect_result) Then detect_result="无" Else detect_result=detect_result&"%"
				detect_report=rsDetect("DETECT_REPORT").Value
%><li><%=index%>.检测时间：<%=detect_time%>，查重结果：<%=detect_result%>
<br/><a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=8&hash=<%=hash%>" target="_blank">送检论文</a><%
				If Not IsNull(detect_report) Then %>
&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=12&hash=<%=hash%>" target="_blank">检测报告</a><%
				End If
%></li><%
				index=index+1
				rsDetect.MoveNext()
			Loop
%></ul><%
		End If
%></td></tr><%
	End If
	If review_status>=rsDetectThesisUploaded And Len(thesis_file_review) Then %>
<tr><td>送审论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=9" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsAgreeReview Then
		If Not IsNull(review_app) Then %>
<tr><td>送审申请表：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=13" target="_blank" >点击下载</a></td></tr><%
		End If
		' 根据评阅书显示设置决定是否显示文件
		If numReviewed And bIsReviewVisible Then %>
<tr><td>论文评阅书：&emsp;&emsp;<%
			If Len(review_file(0)) Then
%>1.<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=14" target="_blank">点击下载</a>（返回于&nbsp;<%=review_time(0)%>）&emsp;<%
			End If
			If Len(review_file(1)) Then
%>2.<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=15" target="_blank">点击下载</a>（返回于&nbsp;<%=review_time(1)%>）<%
			End If
%></td></tr><%
		End If
	End If
	If review_status>=rsModifyThesisUploaded And Len(thesis_file_modified) Then %>
<tr><td>答辩论文：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=10" target="_blank">点击下载</a></td></tr><%
	End If
	If review_status>=rsFinalThesisUploaded And Len(thesis_file_final) Then %>
<tr><td>定稿论文：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=11" target="_blank">点击下载</a>&emsp;<a href="#" onclick="return rollback(<%=thesisID%>,0,5)">撤销</a></td></tr><%
	End If
	If task_progress>=tpTbl4Uploaded And Len(table_file(4)) Then %>
<tr><td>答辩审批材料：&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=7" target="_blank">点击下载</a></td></tr><%
	End If
%><tr><td height="10"></td></tr><%
	If numReviewed And bIsReviewVisible Then %>
<tr><td>对学位论文的总体评价&nbsp;1：<%=reviewLevelRadios("reviewlevel1",reviewfile_type,review_level(0))%></td></tr>
<tr><td>对学位论文的总体评价&nbsp;2：<%=reviewLevelRadios("reviewlevel2",reviewfile_type,review_level(1))%></td></tr>
<tr><td>评审结果&nbsp;1：&emsp;&emsp;<%=reviewResultList("reviewresult",review_result(0),false)%>&emsp;<span class="tip">(A→同意答辩；B→需做适当修改；C→需做重大修改；D→不同意答辩；E→尚未返回)</span></td></tr>
<tr><td>评审结果&nbsp;2：&emsp;&emsp;<%=reviewResultList("reviewresult",review_result(1),false)%>&emsp;<span class="tip">(A→同意答辩；B→需做适当修改；C→需做重大修改；D→不同意答辩；E→尚未返回)</span></td></tr>
<tr><td>处理意见：&emsp;&emsp;&emsp;<%=finalResultList("reviewresult",review_result(2),false)%></td></tr><%
	End If
	If Not IsNull(rs("TASK_EVAL")) Then %>
<tr><td>导师对表格的审核意见：&emsp;<%=toPlainString(rs("TASK_EVAL"))%></td></tr><%
	End If
	If review_status>=rsNotAgreeDetect Then %>
<tr><td>导师送检意见：&emsp;<%=toPlainString(rs("DETECT_APP_EVAL"))%></td></tr><%
	End If
	If review_status=rsAgreeDetect Or review_status>=rsNotAgreeReview Then %>
<tr><td>导师送审意见（<%=rs("SUBMIT_REVIEW_TIME")%>）：&emsp;<%=toPlainString(rs("REVIEW_APP_EVAL"))%></td></tr><%
	End If
	If review_status>=rsModifyUnpassed Then %>
<tr><td><%=tutor_modify_eval_title%>：<%=toPlainString(rs("TUTOR_MODIFY_EVAL"))%></td></tr><%
	End If
	If review_status>=rsDefenceEval Then %>
<tr><td>答辩委员会修改意见：<br/><%=toPlainString(rs("DEFENCE_EVAL"))%></td></tr><%
	End If
	If review_status>=rsInstructEval Then %>
<tr><td>学院学位评定分委员会修改意见：<br/><%=toPlainString(rs("INSTRUCT_MODIFY_EVAL"))%></td></tr><%
	End If
	If Not IsNull(rs("DEFENCE_MEMBER")) Then
		Dim defence_members,defence_memo
		defence_members=Split(rs("DEFENCE_MEMBER"),"|")
		defence_memo=rs("MEMO")
		If IsNull(defence_memo) Then defence_memo="-" %>
<tr><td>答辩安排：
<div id="defenceplan"><table cellspacing="0" cellpadding="1">
<thead><tr style="font-weight:bold"><td width="100"><p>答辩时间</p></td><td width="100"><p>答辩地点</p></td>
<td width="60"><p>答辩主席</p></td><td width="100"><p>答辩委员</p></td><td width="60"><p>答辩秘书</p></td><td><p>答辩委员工作单位</p></td></tr></thead>
<tbody><tr><td><p><%=rs("DEFENCE_TIME")%></p></td><td><p><%=rs("DEFENCE_PLACE")%></p></td>
<td><p><%=defence_members(0)%></p></td><td><p><%=defence_members(1)%></p></td><td><p><%=defence_members(2)%></p></td>
<td><p><%=toPlainString(defence_memo)%></p></td></tbody></table></div></td></tr><%
	End If
	If defence_result<>0 Then %>
<tr><td>答辩成绩：&emsp;&emsp;&emsp;&emsp;&emsp;<%=defenceResultList("defenceresult",defence_result)%></td></tr><%
	End If
	If Not IsNull(grant_degree) Then %>
<tr><td>是否同意授予学位：&emsp;<%=grantDegreeList("grantdegree",grant_degree)%></td></tr><%
	End If %>
<tr class="trbuttons">
<td colspan="3"><p align="center"><%
	If nSystemStatus=2 Then
		Select Case opr
		Case 1,2,3,4 %>
<input type="button" id="unpass" name="btnsubmit" value="<%=arrCaptionNotAgree(opr)%>" />&emsp;
<input type="button" id="pass" name="btnsubmit" value="<%=arrCaptionAgree(opr)%>" />&emsp;<%
		Case 5 %>
<input type="button" id="unpass" name="btnsubmit" value="不同意该生论文查重、送审" />&emsp;
<input type="button" id="pass" name="btnsubmit" value="同意该生论文查重、查重结果低于10%系统自动匹配送审" />&emsp;<%
		Case 6 %>
<input type="button" id="unpass" name="btnsubmit" value="不同意送审" />&emsp;
<input type="button" id="pass" name="btnsubmit" value="同意送审" />&emsp;<%
		Case 7
			If bIsReviewVisible Then %>
<input type="button" id="btnsubmit" name="btnsubmit" value="确认评阅结果" />&emsp;<%
			End If
		Case 8 %>
<input type="button" id="unpass" name="btnsubmit" value="不同意论文修改" />&emsp;
<input type="button" id="pass" name="btnsubmit" value="确认修改，同意答辩" />&emsp;<%
		Case 9
			If Not IsEmpty(reviewer) Then %>
<input type="button" id="btnsubmit" name="btnsubmit" value="评阅该论文" />&emsp;<%
			End If
		End Select
	End If
%><input type="button" value="关 闭" onclick="closeWindow()" />
</p></td></tr></table>
<input type="hidden" name="stuid" value="<%=rs("STU_ID")%>" />
<input type="hidden" name="opr" value="<%=opr%>" />
<input type="hidden" id="submittype" name="submittype" />
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<table class="tblform" width="800" cellspacing=1 cellpadding=3>
<tr style="background-color: #cccccc"><td><p>论文检测结果及论文评审结果说明：</p>
<%=getNoticeText(rs("TEACHTYPE_ID"),"review_result_desc")%>
</td></tr></table></center>
<form id="ret" name="ret" action="thesisList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
</body><script type="text/javascript">
	var btnsubmit=document.getElementsByName("btnsubmit");
	var arrActionUrl=["<%=actionUrl1%>","<%=actionUrl2%>"];
	if(btnsubmit) {
		for(var i=0;i<btnsubmit.length;i++) {
			btnsubmit.item(i).action=arrActionUrl[i];
			btnsubmit.item(i).onclick=function() {
				this.value="正在提交，请稍候……";
				this.disabled=true;
				this.form.submittype.value=this.id;
				this.form.action=this.action;
				this.form.submit();
			}
			btnsubmit.item(i).disabled=false;
		}
	}
</script></html><%
Case 2	' 填写评语页面
	opr=Request.Form("opr")
	submittype=Request.Form("submittype")
	isunpass=submittype="unpass"
	Select Case opr
	Case 1,2,3,4
		If isunpass Then
			operation_name="您审核不通过"&arrTable(opr)&"，请填写审核意见"
		ElseIf opr=4 Then
			operation_name="您审核通过了"&arrTable(opr)&"，请填写指导意见"
		End If
	Case 5
		If isunpass Then
			operation_name="您不同意论文检测和送审，请填写审核意见"
		Else
			operation_name="您同意了论文检测和送审，请填写送审评语"
		End If
	Case 6
		If isunpass Then
			operation_name="您不同意论文送审，请填写审核意见"
		Else
			operation_name="您同意了论文送审，请填写送审评语"
		End If
	Case 8
		If isunpass Then
			operation_name="您不同意论文修改，请填写意见"
		Else
			operation_name="您同意该生参加论文答辩，请填写意见"
		End If
	End Select
	tutor_duty_name=getProDutyNameOf(tutor_id)
	eval_text=rs("TUTOR_MODIFY_EVAL")
	If IsNull(eval_text) Then
		eval_text=rs("TUTOR_REVIEW_EVAL")
	Else
		eval_text="上一环节（答辩论文）审核评语：<br/>"&eval_text
	End If
	If IsNull(eval_text) Then
		eval_text=rs("REVIEW_APP_EVAL")
	Else
		eval_text="上一环节（评阅意见）审核评语：<br/>"&eval_text
	End If
	If IsNull(eval_text) Then
		eval_text=rs("DETECT_APP_EVAL")
		If Not IsNull(eval_text) Then
			eval_text="上一环节（送检论文）审核评语：<br/>"&eval_text
		End If
	Else
		eval_text="上一环节（送审论文）审核评语：<br/>"&eval_text
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/tutor.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<script src="../scripts/drafting.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>填写审核意见</title>
<style type="text/css" />
	input[type="text"] { background:none;border-top:0;border-left:0;border-right:0;border-bottom:1px solid dimgray }
</style>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b><%=operation_name%></b></font>
<form id="fmOper" action="updateThesis.asp?tid=<%=thesisID%>" method="post" style="margin-top:0;padding-top:10px">
<table class="tblform" width="800" cellspacing="1" cellpadding="3">
<tr><td>作者姓名：<input type="text" class="txt" name="author" value="<%=rs("STU_NAME")%>" readonly /></td>
<td>学号：<input type="text" class="txt" name="stuno" value="<%=rs("STU_NO")%>" readonly /></td>
<td>导师姓名、职称：<input type="text" class="txt" name="tutorinfo" value="<%=Session("Tname")%>&nbsp;<%=tutor_duty_name%>" readonly /></td></tr>
<tr><td colspan=2>申请学位专业名称：<input type="text" class="txt" name="speciality" size="50" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td>
<td>学院名称：<input type="text" class="txt" name="faculty" value="工商管理学院" readonly /></td></tr>
<tr><td colspan=3>学位论文题目：<input type="text" class="txt" name="subject" size="70" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr><%
	Select Case opr
	Case 1,2,3,4 ' 填写表格审核意见页面
		Select Case opr
		Case 1
			If Not IsNull(rs("TABLE_FILE1")) Then %>
<tr><td colspan=3>开题报告表：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=1" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE1")) Then %>
<tr><td colspan=3>开题论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=2" target="_blank">点击下载</a></td></tr><%
			End If
		Case 2
			If Not IsNull(rs("TABLE_FILE2")) Then %>
<tr><td colspan=3>中期检查表：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=3" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE2")) Then %>
<tr><td colspan=3>中期论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=4" target="_blank">点击下载</a></td></tr><%
			End If
		Case 3
			If Not IsNull(rs("TABLE_FILE3")) Then %>
<tr><td colspan=3>预答辩申请表：&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=5" target="_blank">点击下载</a></td></tr><%
			End If
			If Not IsNull(rs("TBL_THESIS_FILE3")) Then %>
<tr><td colspan=3>预答辩论文：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=6" target="_blank">点击下载</a></td></tr><%
			End If
		Case 4 %>
<tr><td colspan=3>答辩审批材料：&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=7" target="_blank">点击下载</a></td></tr><%
		End Select
		If isunpass Then
			eval_text_name=arrTable(opr)&"审核意见（200-2000字）："
		ElseIf opr=4 Then
			eval_text_name="校内指导教师意见（包括对申请人的学习情况、思想表现及论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见，200-2000字）"
		End If %>
<tr><td colspan=3><%=eval_text_name%><span id="eval_text_tip"></span></td></tr>
<tr><td colspan=3><textarea name="eval_text" rows="15" style="width:100%"><%
	If opr=4 Then
		Response.Write eval_text
	End If
%></textarea></td></tr><%
	Case 5 ' 填写论文送检送审意见页面 %>
<tr><td colspan="3">送检论文：<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=8" target="_blank">点击下载</a>&emsp;送审论文：<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=9" target="_blank">点击下载</a></td></tr><%
		If Not rsDetect.EOF Then
%><tr><td colspan="3">上次论文检测记录：<%
			detect_time=rsDetect("DETECT_TIME").Value
			If IsNull(detect_time) Then detect_time="无"
			detect_result=rsDetect("RESULT").Value
			If IsNull(detect_result) Then detect_result="无" Else detect_result=detect_result&"%"
			detect_report=rsDetect("DETECT_REPORT").Value
%>检测时间：<%=detect_time%>，查重结果：<%=detect_result%><%
			If Not IsNull(detect_report) Then %>
<br/><a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=8&time=<%=detect_time%>" target="_blank">送检论文</a>
<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=12&time=<%=detect_time%>" target="_blank">检测报告</a><%
			End If
%></td></tr><%
		End If
%><tr><td colspan=3>导师对论文的意见<span class="eval_notice">（200-2000字，包含选题意义；文献资料的掌握；数据、材料的收集、论证、结论是否合理；基本论点、结论和建议有无理论意义和实践价值等）</span>：<br/><span class="tip">提示：复制比低于10%的学员，系统会自动匹配进行论文盲审。复制比高于10%的学员，由中心统一进行二次查重，二次查重所产生的费用由学员本人缴纳。</span>&emsp;<span id="eval_text_tip"></span></td></tr>
<tr><td colspan=3><textarea name="eval_text" rows="15" style="width:100%"></textarea></td></tr><%
	Case 6 ' 填写导师送审评语页面 %>
<tr><td colspan=3>送审论文：<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=9" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>导师对学位论文的评语<span class="eval_notice">（请阅读论文后填写，200-2000字）</span>：<span id="eval_text_tip"></span><br/>
送审评语的基本内容参考：<br/><%=getNoticeText(rs("TEACHTYPE_ID"),"review_eval_reference")%></td></tr>
<tr><td colspan=3><textarea name="eval_text" rows="10" style="width:100%"></textarea><br/></td></tr><%
		If Not isunpass Then %>
<tr><td colspan=3 style="padding:0"><table class="tblform" width="100%" cellspacing="1" cellpadding="3">
<tr><td width="100" align="center">作者承诺</td>
<td><p>&emsp;&emsp;1．该学位论文为公开学位论文，其中不涉及国家秘密项目和其它不宜公开的内容，否则将由本人承担因学位论文涉密造成的损失和相关的法律责任；<br/>
&emsp;&emsp;2．该学位论文是本人在导师的指导下独立进行研究所取得的研究成果，不存在学术不端行为。</p>
<p align="right">作者签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<%=toDateTime(Now(),1)%></p></td></tr>
<tr><td align="center">指导教师<br/>意见</td>
<td><p><span style="font-size:15pt">■</span>&nbsp;同意送审<br/><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
<p align="right">指导教师签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=toDateTime(Now(),1)%></span></p></td></tr>
<tr><td align="center"></td>
<td><p><span style="font-size:15pt">□</span>&nbsp;同意送审<br/><span style="font-size:15pt">□</span>&nbsp;不同意送审</p>
<p align="right">主管院领导签名：&emsp;&emsp;&emsp;&emsp;&nbsp;日期：<span style="visibility:hidden"><%=toDateTime(Now(),1)%></span></p></td></tr>
<tr><td align="center">备注</td>
<td><p>经图书馆检测，学位论文文字复制比&nbsp;<span style="text-decoration:underline"><%=reproduct_ratio%>%</span><input type="hidden" name="reproduct_ratio" size="10" value="<%=reproduct_ratio%>" /></p></td></tr></table></td></tr><%
		End If
	Case 8 ' 填写修改意见页面 %>
<tr><td colspan=3>答辩论文：<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=10" target="_blank">点击下载</a></td></tr>
<tr><td colspan=3>导师对学位论文的评语<span class="eval_notice">（此意见将嵌入学生《学位论文答辩及授予学位审批材料》中，并进入学籍档案，意见需包含对学生的业务学习、思想表现及学位论文的学术评语，科研工作能力和完成科研工作情况，以及是否同意申请学位论文答辩的意见，200-2000字）</span>：<span id="eval_text_tip"></span></td></tr>
<tr><td colspan=3><textarea name="eval_text" rows="10" style="width:100%"></textarea><br/></td></tr><%
	End Select
	If Not IsNull(eval_text) Then %>
<tr><td colspan=3><%=eval_text%></td></tr><%
	End If %>
<tr class="trbuttons">
<td colspan="3"><p align="center"><input type="button" id="btnsavedraft" value="保存草稿" />&emsp;
<input type="button" id="btnloaddraft" value="读取草稿" />&emsp;
<input type="button" id="btnsubmit" name="btnsubmit" value="提 交" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="关 闭" onclick="closeWindow()" />
</p></td></tr></table>
<input type="hidden" name="opr" value="<%=opr%>" />
<input type="hidden" name="submittype" value="<%=submittype%>" />
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></center>
<form id="ret" name="ret" action="thesisList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form></body>
<script type="text/javascript">
	$(document).ready(function(){
		$('[name="eval_text"]').keyup(function(){checkLength(this,2000)});
		if($('#btnsubmit').size()>0) {
			$('#btnsubmit').click(function() {
				if(confirm("提交后将不能再更改信息，确定要提交吗？")) {
					$(this).val("正在提交，请稍候……")
								 .attr('disabled',true);
					this.form.submit();
				}
			}).attr('disabled',false);
		}
		$('#btnsavedraft').click(function() {
			saveAsDraft(<%=thesisID%>);
		});
		verifyDraft(<%=thesisID%>);
		// 每30秒对草稿进行自动保存
		setInterval('saveAsDraft(<%=thesisID%>,true)',30000);
	});
</script></html><%
End Select
CloseRs rsDetect
CloseRs rs
CloseConn conn
%>