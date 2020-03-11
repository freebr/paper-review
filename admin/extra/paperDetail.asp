<!--#include file="../../inc/global.inc"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../../error.asp?timeout")
paper_id=Request.Form("tid")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
reviewer=Request.Form("rev")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")

If Len(paper_id)=0 And Len(stuname)=0 Then
	bError=True
	errdesc="参数无效。"
Else
	Connect conn
	If Len(paper_id) Then
		sql="SELECT * FROM ViewDissertations WHERE ID="&paper_id
	Else
		sql="SELECT * FROM ViewDissertations WHERE STU_NAME="&toSqlString(stuname)
	End If
	GetRecordSet conn,rs,sql,count
	If count=0 Then
		bError=True
		errdesc="数据库没有该论文记录，或您未受邀评阅该论文！"
	End If
End If
If bError Then
	CloseRs rs
	CloseConn conn
	showErrorPage errdesc, "提示"
End If

Dim review_status
Dim review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)
paper_id=rs("ID")
author_stu_type=rs("TEACHTYPE_ID")
If author_stu_type=5 Or author_stu_type=6 Then
	reviewfile_type=2
Else
	reviewfile_type=1
End If
eval_text=rs("REVIEWER_EVAL"&(reviewer+1))
review_app=rs("REVIEW_APP")
review_status=rs("REVIEW_STATUS")
If Not IsNull(rs("THESIS_FILE2")) Then
	thesis_file_review=rs("THESIS_FILE2")
End If
If Not IsNull(rs("REVIEW_RESULT")) Then
	arr=Split(rs("REVIEW_RESULT"),",")
	For i=0 To UBound(arr)
		review_result(i)=Int(arr(i))
	Next
End If
If Not IsNull(rs("REVIEWER_EVAL_TIME")) Then
	arr2=Split(rs("REVIEWER_MASTER_LEVEL"),",")
	arr3=Array(rs("ReviewFile1"),rs("ReviewFile2"))
	arr4=Split(rs("REVIEWER_EVAL_TIME"),",")
	arr5=Split(rs("REVIEW_LEVEL"),",")
	For i=0 To 1
		reviewer_master_level(i)=Int(arr2(i))
		review_file(i)=arr3(i)
		review_time(i)=arr4(i)
		review_level(i)=Int(arr5(i))
	Next
End If
If Len(review_time(reviewer))=0 Then
	stat="专家尚未评阅"
Else
	stat="专家已评阅"
End If

view_name = "paperDetail_review_"&paper_id
' 获取视图状态
view_state = getViewState(Session("TId"),usertypeExpert,view_name)

tutor_duty_name=getProDutyNameOf(rs("TUTOR_ID"))
If reviewfile_type=2 Then
	loadReviewScoringInfo rs("REVIEW_TYPE"),code_scoring,code_power1,code_power2
End If
If author_stu_type=5 Then
	correlation_level_name="学位论文内容与申请学位领域的相关性："
Else
	correlation_level_name="学位论文内容与申请学位专业的相关性："
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>以【<%=rs("EXPERT_NAME"&(reviewer+1))%>】的身份评阅论文</title>
<% useStylesheet "admin", "jeasyui" %>
<% useScript "jquery", "jeasyui", "common", "paper", "review", "viewState" %>
</head>
<body>
<center><div><font size=4><b>以【<%=rs("EXPERT_NAME"&(reviewer+1))%>】的身份评阅论文</b></font>
<form id="fmReview" action="doReview.asp?tid=<%=paper_id%>&rev=<%=reviewer%>" method="post" style="margin-top:0;padding-top:10px">
<table class="form buttons"><tr><td>
<input type="button" id="btnsavedraft" value="保存草稿" />
<input type="button" id="btnloaddraft" value="读取草稿" />
</td></tr></table>
<table class="form"><tr><td>
<div class="fields">
	<div>申请学位专业名称：<input type="text" class="txt full-width" name="speciality" size="25" value="<%=rs("SPECIALITY_NAME")%>" readonly /></div>
	<div>研究方向：<input type="text" class="txt full-width" name="researchway" size="25" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></div>
	<div>学院名称：<input type="text" class="txt full-width" name="faculty" value="工商管理学院" readonly /></div>
</div>
<div class="fields">
	<div>学位论文题目：<input type="text" class="txt full-width" name="subject" size="70" value="<%=rs("THESIS_SUBJECT")%>" readonly /></div>
	<div>送审论文：<a class="resc" href="fetchDocument.asp?tid=<%=paper_id%>&type=1" target="_blank">点击下载</a></div>
</div>
<div class="fields">
	<div>对本论文涉及内容的熟悉程度：<%=masterLevelRadios("master_level",reviewer_master_level(reviewer))%></div>
</div>
<div class="fields">评阅专家对论文的学术评语<span class="eval_notice">（包括选题意义；文献资料的掌握；数据、材料的收集、论证、结论是否合理；基本论点、结论和建议有无理论意义和实践价值；论文的不足之处和建议等，200-2000字）</span>：</div>
<div class="fields"><textarea name="eval_text" rows="10" style="width:100%">
<%=eval_text%>
</textarea></div>
<div class="fields"><span id="eval_text_tip"></span></div>
</td></tr></table><%
		Dim strJsArrRemarkStd
		Select Case author_stu_type
			Case 5
				strJsArrRemarkStd="[{'name':'优秀','min':85},{'name':'良好','min':70},{'name':'合格','min':60},{'name':'不合格','min':0}]"
%>
<fieldset>
	<legend><%=rs("TEACHTYPE_NAME")%>学位论文评价指标</legend>
	<p>说明：请评审专家在各二级指标得分空格处按百分制打分，系统将自动生成各一级指标得分并最后汇总计算出总分。</p>
	<table class="form">
		<tr><td width="100" align="center">一级指标</td><td align="center">二级指标</td><td align="center">主要观测点</td><td align="center">参考权重</td><td align="center">得分（百分制）</td></tr>
		<%=code_scoring%>
		<tr><td align="center" colspan="3">加权总分</td><td align="center" colspan="2"><span id="total_score"></span></td></tr>
		<tr><td align="center" rowspan="2">对学位论文的总体评价</td><td align="center" colspan="2"><span id="review_level_text">&nbsp;</span></td></tr>
		<tr><td colspan="2"><p>优秀：总分≥85；良好：84≥总分≥70；合格：69≥总分≥60；不合格：总分≤59。<input type="hidden" name="review_level" /></p></td></tr>
	</table>
</fieldset><%
			Case 6
				strJsArrRemarkStd="[{'name':'优秀','min':90},{'name':'良好','min':75},{'name':'合格','min':60},{'name':'不合格','min':0}]"
%>
<fieldset>
	<legend><%=rs("TEACHTYPE_NAME")%>学位论文评价指标</legend>
	<p>说明：请评审专家在各二级指标得分空格处按百分制打分，系统将自动生成各一级指标得分并最后汇总计算出总分。</p>
	<table class="form">
		<tr><td width="6s0" align="center">一级指标</td><td align="center">二级指标</td><td width="350" align="center">评分标准（优秀：≥90；良好：89-75；合格：74-60；不合格：≤59）</td><td align="center">一级指标得分</td></tr>
		<%=code_scoring%><tr><td align="center">总分</td><td align="center" colspan="3"><span id="total_score"></span></td></tr>
		<tr><td align="center" rowspan="2">对学位论文的总体评价</td><td align="center" colspan="3"><span id="review_level_text">&nbsp;</span></td></tr>
		<tr><td colspan="3"><p>优秀：≥90；良好：89-75；合格：74-60；不合格：≤59。<input type="hidden" name="review_level" /></p></td></tr>
	</table>
</fieldset><%
			Case Else
%>
<table class="form">
	<tr><td align="center">对学位论文的总体评价</td><td align="center" colspan="2"><%=reviewLevelRadios("review_level",1,review_level(reviewer))%></td></tr>
</table><%
		End Select %>
<table class="form"><tr><td>
	<div class="fields"><div><%=correlation_level_name%><%=correlationLevelRadios("correlation_level",1)%></div></div>
	<div class="fields"><div>是否同意举行论文答辩：<%=reviewResultRadios("review_result",review_result(reviewer))%></div></div>
</td></tr></table>
<table class="form buttons"><tr><td>
<input type="submit" value="提 交" />
<input type="button" value="返 回" onclick="history.go(-1)" />
<input type="button" value="关 闭" onclick="closeWindow()" />
</td></tr></table>
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></div></center>
<form id="ret" name="ret" action="paperList.asp" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	$(document).ready(function() {
		$('[name="eval_text"]').keyup(function(){checkLength(this,2000)});<%
		If reviewfile_type=2 Then %>
		if ($('#total_score').size()>0) {
			this.powers={'power1':<%=code_power1%>,'power2':<%=code_power2%>};
			this.remarkStd=<%=strJsArrRemarkStd%>;
			addScoreEventListener();
			showTotalScore();
		}<%
		End If %>
		initViewState($("form"), {
			user_id: <%=Session("TId")%>,
			user_type: <%=usertypeExpert%>,
			view_name: "<%=view_name%>",
			view_state: <%=isNullString(view_state, "null")%>
		}, showTotalScore);
		$('form').submit(function() {<%
		If reviewfile_type=2 Then %>
			if($('[name="review_level"]').val()==4)
				if(!confirm("检测到您给出的分数过低，请确认是否对每项得分点按百分制打分。"))
					return false;<%
		End If %>
			if(!confirm("确定要提交吗？")) return false;
			$(':submit').val("正在提交，请稍候……").attr('disabled',true);
		});
		$(':submit').attr('disabled',false);
		// 每30秒保存一次草稿
		setInterval("$('#btnsavedraft').click()",30000);
	});
</script></body></html><%
CloseRs rs
CloseConn conn
%>