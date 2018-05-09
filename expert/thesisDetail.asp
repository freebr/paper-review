<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
If Not hasPrivilege(Session("Twriteprivileges"),"I10") Then Response.Redirect("../error.asp?no-privilege")
curstep=Request.QueryString("step")
thesisID=Request.QueryString("tid")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
spec_id=Request.Form("In_SPECIALITY_ID2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
finishReview=Request.Form("finishReview")
willUpdateList=LCase(finishReview="1")

Connect conn
sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID&" AND "&Session("Tid")&" IN (REVIEWER1,REVIEWER2)"
GetRecordSet conn,rs,sql,result
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Then
	bError=True
	errdesc="参数无效。"
ElseIf result=0 Then
	bError=True
	errdesc="数据库没有该论文记录，或您未受邀评阅该论文！"
ElseIf Not checkIfProfileFilledIn() Then
	bError=True
	errdesc="在开始评阅前，您需要完善个人信息，<a href=""profile.asp"">请点击这里编辑。</a>"
End If
If bError Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /><script src="../scripts/thesis.js" type="text/javascript"></script><title>查看论文信息</title></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="关 闭" onclick="closeWindow()" /></div></center></body></html><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If

Dim author_stu_type,reviewer,review_status
Dim review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)
Dim formAction

author_stu_type=rs("TEACHTYPE_ID")
Select Case author_stu_type
Case 5,6,9
	formAction="?tid="&thesisID&"&step=2"
Case Else
	formAction="?tid="&thesisID&"&step=3"
End Select

If rs("REVIEWER1")=Session("Tid") Then
	reviewer=0
Else
	reviewer=1
End If
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
	arr3=Split(rs("REVIEW_FILE"),",")
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
	stat="您尚未评阅"
Else
	stat="您已评阅"
End If
Select Case curstep
Case vbNullString	' 论文详情页面
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/tutor.css" rel="stylesheet" type="text/css" />
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>查看论文信息</title>
</head>
<body class="exp">
<center><div class="content"><font size=4><b>专业硕士论文详情<br />论文当前状态：【<%=stat%>】</b></font><%
	If Len(review_time(reviewer))=0 And nSystemStatus<>2 Then
%><p><span class="tip">温馨提示：评阅系统的开放时间为<%=toDateTime(startdate,1)%>至<%=toDateTime(enddate,1)%>，当前不在开放时间内，不能评阅论文。</span></p><%
	End If %>
<form id="fmDetail" action="<%=formAction%>" method="post">
<table class="tblform" width="800" align="center" cellspacing="1" cellpadding="3">
<tr><td>论文题目：&emsp;&emsp;&emsp;<input type="text" class="txt" name="subject" size="95%" value="<%=rs("THESIS_SUBJECT")%>" readonly /></td></tr>
<tr><td>（英文）：&emsp;&emsp;&emsp;<input type="text" class="txt" name="subject_en" size="85%" value="<%=rs("THESIS_SUBJECT_EN")%>" readonly /></td></tr>
<tr><td>学位类别：&emsp;&emsp;&emsp;<input type="text" class="txt" name="degreename" size="95%" value="<%=rs("TEACHTYPE_NAME")%>" readonly /></td></tr><%
	If reviewfile_type=2 Then %>
<tr><td>领域名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="speciality" size="95%" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td></tr><%
	End If %>
<tr><td>研究方向：&emsp;&emsp;&emsp;<input type="text" class="txt" name="researchway_name" size="95%" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td></tr>
<tr><td>论文关键词：&emsp;&emsp;<input type="text" class="txt" name="keywords_ch" size=85%" value="<%=rs("KEYWORDS")%>" readonly /></td></tr>
<tr><td>（英文）：&emsp;&emsp;&emsp;<input type="text" class="txt" name="keywords_en" size="85%" value="<%=rs("KEYWORDS_EN")%>" readonly /></td></tr>
<tr><td>院系名称：&emsp;&emsp;&emsp;<input type="text" class="txt" name="faculty" size="95%" value="工商管理学院" readonly /></td></tr><%
	If Not IsNull(rs("THESIS_FORM")) And Len(rs("THESIS_FORM")) Then %>
<tr><td>论文形式：&emsp;&emsp;&emsp;<input type="text" class="txt" name="thesisform" size="95%" value="<%=rs("THESIS_FORM")%>" readonly /></td></tr><%
	End If %>
<tr><td>送审论文：&emsp;&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=1" target="_blank">点击下载</a></td></tr><%
	If Len(review_time(reviewer)) Then
%><tr><td height="10"></td></tr><%
		If Len(review_file(reviewer)) Then
%><tr><td>您已于&nbsp;<%=toDateTime(review_time(reviewer),1)%>&nbsp;<%=toDateTime(review_time(reviewer),4)%>&nbsp;评阅该论文</td></tr>
<tr><td>论文评阅书：&emsp;&emsp;<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=<%=2+reviewer%>" target="_blank">点击下载</a></td></tr><%
		End If %>
<tr><td>您对本论文涉及内容的熟悉程度：<%=masterLevelRadios("masterlevel",reviewer_master_level(reviewer))%></td></tr>
<tr><td>对学位论文的总体评价：<%=reviewLevelRadios("reviewlevel",reviewfile_type,review_level(reviewer))%></td></tr>
<tr><td>您的评审结果：&emsp;<%=reviewResultList("reviewresult",review_result(reviewer),false)%>&emsp;<span class="tip">(A→同意答辩；B→需做适当修改；C→需做重大修改；D→不同意答辩；E→尚未返回)</span></td></tr><%
	End If %>
<tr class="trbuttons">
<td colspan="3"><p align="center"><%
	If nSystemStatus=2 Then
		btnsubmittext=""
		If Len(review_time(reviewer))=0 Then
			btnsubmittext="评阅该论文"
		ElseIf review_status=rsMatchExpert Or review_status=rsReviewed Then
			btnsubmittext="重新评阅该论文"
		End If
		If Len(btnsubmittext) Then
%><input type="button" id="btnsubmit" name="btnsubmit" value="<%=btnsubmittext%>" />&emsp;<%
		End If
	End If
%><input type="button" value="关 闭" onclick="closeWindow(<%=willUpdateList%>)" />
</p></td></tr></table>
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
</div></center>
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
	if(document.all.btnsubmit) {
		document.all.btnsubmit.onclick=function() {
			this.value="正在提交，请稍候……";
			this.disabled=true;
			this.form.submit();
		}
		document.all.btnsubmit.disabled=false;
	}
</script></html><%
Case 2	' 评阅须知

	Dim noticeName,pageCount,arrNoticeFile
	Select Case author_stu_type
	Case 5
		noticeName="华南理工大学工程硕士学位论文撰写要求（试行）"
		arrNoticeFile=Array("","me1.png","me2.png","me3.png","me4.png")
	Case 6
		noticeName="华南理工大学工商管理硕士学位论文撰写要求（试行）"
		arrNoticeFile=Array("","mba1.png","mba2.png","mba3.png")
	Case 9
		noticeName="华南理工大学会计硕士学位论文撰写要求（试行）"
		arrNoticeFile=Array("","mpacc1.png","mpacc2.png")
	Case Else
		arrNoticeFile=Array("")
	End Select
	pageCount=UBound(arrNoticeFile)
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/tutor.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>浏览评阅须知</title>
</head>
<body class="exp">
<center><div class="content"><font size=4><b>评阅须知：在开始评阅前，请您仔细阅读下面的《<%=noticeName%>》</b></font>
<form id="fmDetail" action="?tid=<%=thesisID%>&step=3" method="post">
<table class="tblform" width="800" align="center" cellspacing="1" cellpadding="3">
<tr><td><p align="center"><anchor id="top" />
<span class="tip">共&nbsp;<%=pageCount%>&nbsp;页，当前是第&nbsp;<span id="numPage"></span>&nbsp;页</span></p></td></tr>
<tr><td><div id="noticeContent"><p align="center">
<%
	For i=1 To UBound(arrNoticeFile)
%><img id="page<%=i%>" src="images/<%=arrNoticeFile(i)%>" style="display:none" /><%
	Next
%></p></div>
<p align="center"><button id="btnPrev">&lt;&lt;上一页</button>
<button id="btnNext">下一页&gt;&gt;</button></p></td></tr>
<tr class="trbuttons">
<td><p align="center">
<input type="button" id="btnsubmit" name="btnsubmit" value="我知道了，开始评阅论文" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="关 闭" onclick="closeWindow()" />
</p></td></tr></table>
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></div></center>
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
	$(document).ready(function(){
		var numPage,pageCount=<%=pageCount%>;
		function prevPage() {
			if(numPage<=1) return false;
			displayPage(numPage-1);
			return false;
		}
		function nextPage() {
			if(numPage>=pageCount) return false;
			displayPage(numPage+1);
			return false;
		}
		function displayPage(pg) {
			numPage=pg;
			$('div#noticeContent img').hide();
			$('div#noticeContent img#page'+pg).show();
			$('span#numPage').text(pg.toString());
			$('button#btnPrev').css('visibility',(pg==1)?'hidden':'visible');
			$('button#btnNext').css('visibility',(pg==pageCount)?'hidden':'visible');
			$('body').animate({'scrollTop':$('anchor#top').offset().top},500);
			return;
		}
		$('button#btnPrev,button#btnNext').attr('style','width:40%;height:50px;cursor:pointer');
		$('button#btnPrev').click(prevPage);
		$('button#btnNext').click(nextPage);
		$('div#noticeContent').attr('style','display:block;position:relative;width:100%;height:auto');
		$('input#btnsubmit').click(function() {
			$(this).attr('value',"正在提交，请稍候……")
						 .attr('disabled',true);
			$(this.form).submit();
		}).attr('disabled',false);
		displayPage(1);
	});
</script></html><%
Case 3	' 评阅页面
	tutor_duty_name=getProDutyNameOf(rs("TUTOR_ID"))
	If reviewfile_type=2 Then
		loadReviewScoringInfo rs("REVIEW_TYPE"),scoringtbl_code,power1code,power2code
	End If
	If author_stu_type=5 Then
		correlationtypename="学位论文内容与申请学位领域的相关性："
	Else
		correlationtypename="学位论文内容与申请学位专业的相关性："
	End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/tutor.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<script src="../scripts/review.js" type="text/javascript"></script>
<script src="../scripts/drafting.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>评阅论文</title>
</head>
<body class="exp">
<center><div class="content"><font size=4><b>评阅论文</b></font>
<form id="fmReview" action="doReview.asp?tid=<%=thesisID%>" method="post" style="margin-top:0;padding-top:10px">
<table class="tblform" width="800" cellspacing="1" cellpadding="3">
<tr><td>申请学位专业名称：<input type="text" class="txt" name="speciality" size="25" value="<%=rs("SPECIALITY_NAME")%>" readonly /></td>
<td>研究方向：<input type="text" class="txt" name="researchway" size="25" value="<%=rs("RESEARCHWAY_NAME")%>" readonly /></td>
<td>学院名称：<input type="text" class="txt" name="faculty" value="工商管理学院" readonly /></td></tr>
<tr><td colspan="3">学位论文题目：<input type="text" class="txt" name="subject" size="70" value="<%=rs("THESIS_SUBJECT")%>" readonly />
&emsp;送审论文：<a class="resc" href="fetchfile.asp?tid=<%=thesisID%>&type=1" target="_blank">点击下载</a></td></tr>
<tr><td colspan="3">对本论文涉及内容的熟悉程度：<%=masterLevelRadios("masterlevel",reviewer_master_level(reviewer))%></td></tr>
<tr><td colspan="3">评阅专家对论文的学术评语<span class="eval_notice">（包括选题意义；文献资料的掌握；数据、材料的收集、论证、结论是否合理；基本论点、结论和建议有无理论意义和实践价值；论文的不足之处和建议等，200-2000字）</span>：<span id="eval_text_tip"></span></td></tr>
<tr><td colspan="3"><textarea name="eval_text" rows="10" style="width:100%">
<%=eval_text%></textarea><br/>
<input type="button" id="btnsavedraft" value="保存草稿" />&emsp;
<input type="button" id="btnloaddraft" value="读取草稿" /></td></tr><%
	Dim strJsArrRemarkStd
	Select Case author_stu_type
	Case 5
		strJsArrRemarkStd="[{'name':'优秀','min':85},{'name':'良好','min':70},{'name':'合格','min':60},{'name':'不合格','min':0}]"
	%>
<tr><td align="center" colspan="3" style="padding:0"><p style="font-size:10pt;font-weight:bold"><%=rs("TEACHTYPE_NAME")%>学位论文评价指标</p><table class="tblform" width="100%" cellspacing=1 cellpadding=3>
<tr><td width="20" align="center">一级指标</td><td align="center">二级指标</td><td align="center">主要观测点</td><td align="center">参考权重</td><td align="center">得分（百分制）</td></tr>
<%=scoringtbl_code%><tr><td align="center" colspan="3">加权总分</td><td align="center" colspan="2"><span id="totalscore"></span></td></tr></table></td></tr>
<tr><td align="center" rowspan="2">对学位论文的总体评价</td><td align="center" colspan="2"><span id="reviewleveltext">&nbsp;</span></td></tr>
<tr><td colspan="2"><p>优秀：总分≥85；良好：84≥总分≥70；合格：69≥总分≥60；不合格：总分≤59。<input type="hidden" name="reviewlevel" /></p></td></tr><%
	Case 6
		strJsArrRemarkStd="[{'name':'优秀','min':90},{'name':'良好','min':75},{'name':'合格','min':60},{'name':'不合格','min':0}]"
	%>
<tr><td align="center" colspan="3" style="padding:0"><p style="font-size:10pt;font-weight:bold"><%=rs("TEACHTYPE_NAME")%>学位论文评价指标</p></td></tr>
<tr><td colspan="3"><p>说明：请评审专家在各二级指标得分空格处按百分制打分，系统将自动生成各一级指标得分并最后汇总计算出总分。</p><table class="tblform" width="100%" cellspacing=1 cellpadding=3>
<tr><td width="20" align="center">一级指标</td><td align="center">二级指标</td><td width="350" align="center">评分标准（优秀：≥90；良好：89-75；合格：74-60；不合格：≤59）</td><td align="center">一级指标得分</td></tr>
<%=scoringtbl_code%><tr><td align="center">总分</td><td align="center" colspan="3"><span id="totalscore"></span></td></tr></table></td></tr>
<tr><td align="center" rowspan=2>对学位论文的总体评价</td><td align="center" colspan="3"><span id="reviewleveltext">&nbsp;</span></td></tr>
<tr><td colspan="3"><p>优秀：≥90；良好：89-75；合格：74-60；不合格：≤59。<input type="hidden" name="reviewlevel" /></p></td></tr><%
	Case Else %>
<tr><td align="center">对学位论文的总体评价</td><td align="center" colspan="2"><%=reviewLevelRadios("reviewlevel",1,review_level(reviewer))%></td></tr><%
	End Select %>
<tr><td align="center"><%=correlationtypename%></td><td align="center" colspan="2"><%=correlationTypeRadios("correlationtype",1)%></td></tr>
<tr><td align="center">是否同意举行论文答辩</td><td align="center" colspan="2"><%=reviewResultRadios("reviewresult",review_result(reviewer))%></td></tr>
<tr class="trbuttons">
<td colspan="3"><p align="center"><input type="button" id="btnsubmit" name="btnsubmit" value="提 交" />&emsp;
<input type="button" value="返 回" onclick="history.go(-1)" />&emsp;
<input type="button" value="关 闭" onclick="closeWindow()" />
</p></td></tr></table>
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_SPECIALITY_ID2" value="<%=spec_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form></div></center>
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
		$('[name="eval_text"]').keyup(function(){checkLength(this,2000)});<%
				If reviewfile_type=2 Then %>
		if ($('#totalscore').size()>0) {
			this.powers={'power1':<%=power1code%>,'power2':<%=power2code%>};
			this.remarkStd=<%=strJsArrRemarkStd%>;
			addScoreEventListener();
			showTotalScore();
		}<%
				End If %>
		if($('#btnsubmit').size()>0) {
			$('#btnsubmit').click(function() {<%
				If reviewfile_type=2 Then %>
				if($('[name="reviewlevel"]').val()==4)
					if(!confirm("检测到您给出的分数过低，请确认是否对每项得分点按百分制打分。"))
						return false;<%
				End If %>
				if(confirm("确定要提交吗？")) {
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
CloseRs rs
CloseConn conn
%>