<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")

Dim PubTerm,PageNo,PageSize
sem_info=getCurrentSemester()
period_id=Request.Form("In_PERIOD_ID")
teachtype_id=Request.Form("In_TEACHTYPE_ID")
class_id=Request.Form("In_CLASS_ID")
enter_year=Request.Form("In_ENTER_YEAR")
query_task_progress=Request.Form("In_TASK_PROGRESS")
query_review_status=Request.Form("In_REVIEW_STATUS")
finalFilter=Request.Form("finalFilter")
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"
If Len(period_id) And period_id<>"0" Then
	period_id=Int(period_id)
	PubTerm=PubTerm&" AND PERIOD_ID="&toSqlString(period_id)
Else
	period_id=sem_info(3)
End If
If Len(teachtype_id) And teachtype_id<>"0" Then
	teachtype_id=Int(teachtype_id)
	PubTerm=PubTerm&" AND TEACHTYPE_ID="&toSqlString(teachtype_id)
Else
	teachtype_id=0
End If
If Len(class_id) And class_id<>"0" Then
	class_id=Int(class_id)
	PubTerm=PubTerm&" AND SPECIALITY_ID="&toSqlString(class_id)
Else
	class_id=0
End If
If Len(enter_year) And enter_year<>"0" Then
	enter_year=Int(enter_year)
	PubTerm=PubTerm&" AND ENTER_YEAR="&toSqlString(enter_year)
Else
	enter_year=0
End If
If Len(query_task_progress) And query_task_progress<>"-1" Then
	PubTerm=PubTerm&" AND TASK_PROGRESS="&toSqlString(query_task_progress)
End If
If Len(query_review_status) And query_review_status<>"-1" Then
	PubTerm=PubTerm&" AND REVIEW_STATUS="&toSqlString(query_review_status)
End If

'----------------------PAGE-------------------------
PageNo=""
PageSize=""
If Request.Form("In_PageNo").Count=0 Then
	PageNo=Request.Form("pageNo")
	PageSize=Request.Form("pageSize")
Else
	PageNo=Request.Form("In_PageNo")
	PageSize=Request.Form("In_PageSize")
End If
bShowAll=Request.QueryString="showAll"
If bShowAll Then PageSize=-1
'------------------------------------------------------
Dim arrReviewFileStat
arrReviewFileStat=getReviewFileStatTxtArray()
If Len(PubTerm) Then
	Connect conn
	sql="SELECT ID,THESIS_SUBJECT,STU_NAME,STU_NO,SPECIALITY_NAME,TEACHTYPE_ID,TEACHTYPE_NAME,TUTOR_NAME,REVIEW_STATUS,TASK_PROGRESS,dbo.getThesisStatusText(1,TASK_PROGRESS,1) AS STAT_TEXT1,dbo.getThesisStatusText(2,REVIEW_STATUS,1) AS STAT_TEXT2,REVIEWER1,REVIEWER2,REVIEW_RESULT,REVIEW_FILE_STATUS,"&_
			"CASE WHEN TASK_PROGRESS IN (1,4,7,10) THEN 1 WHEN REVIEW_STATUS=3 THEN 1 WHEN REVIEW_STATUS=7 AND (REVIEWER1 IS NULL OR REVIEWER2 IS NULL) THEN 1 ELSE 0 END AS UNHANDLED FROM VIEW_TEST_THESIS_REVIEW_INFO "&_
			"WHERE VALID=1 "&PubTerm&" ORDER BY UNHANDLED DESC,CLASS_ID DESC,REVIEW_STATUS DESC,TASK_PROGRESS DESC"
	GetRecordSetNoLock conn,rs,sql,result
	If IsEmpty(pageSize) Or Not IsNumeric(pageSize) Then
	  pageSize=60
	Else
		pageSize=CInt(pageSize)
	End If
	If pageSize=-1 Then
		If rs.RecordCount>0 Then rs.PageSize=rs.RecordCount
	Else
	  rs.PageSize=pageSize
	End If
	If IsEmpty(pageNo) Or Not IsNumeric(pageNo) Then
		If rs.PageCount<>0 Then pageNo=1
	Else
		pageNo=CInt(pageNo)
	  If pageNo>rs.PageCount Then
		  If rs.PageCount<>0 Then pageNo=1
		End If
	End If
	If rs.RecordCount>0 Then rs.AbsolutePage=pageNo
End If
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/admin.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.6.3.min.js" type="text/javascript"></script>
<script src="../scripts/query.js" type="text/javascript"></script>
<script src="../scripts/thesis.js" type="text/javascript"></script>
<title>专业论文列表</title>
</head>
<body bgcolor="ghostwhite" onload="return On_Load()">
<center>
<font size=4><b>专业硕士论文列表</b></font>
<table cellspacing="4" cellpadding="0">
<form id="query_nocheck" method="post" onsubmit="if(Chk_Select())return chkField();else return false">
<tr><td>学期&nbsp;<%=semesterList("In_PERIOD_ID",Int(period_id))%></td>
<td><table width="100%" cellspacing="4" cellpadding="0"><%
Dim ArrayList(2,5),k

FormName="query_nocheck"
k=0
ArrayList(k,0)="学位类别"
ArrayList(k,1)="VIEW_CODE_TEACHTYPE"
ArrayList(k,2)="TEACHTYPE_ID"
ArrayList(k,3)="TEACHTYPE_NAME"
ArrayList(k,4)=teachtype_id
ArrayList(k,5)="AND TEACHTYPE_ID IN (5,6,7,9)"

k=1

ArrayList(k,0)="年级"
ArrayList(k,1)="VIEW_SPECIALITY_CLASS"
ArrayList(k,2)="ENTER_YEAR"
ArrayList(k,3)="CAST(ENTER_YEAR AS nvarchar(10))+'级'"
ArrayList(k,4)=enter_year
ArrayList(k,5)="AND VALID=0 AND ENTER_YEAR>=2008"

k=2
ArrayList(k,0)="班级"
ArrayList(k,1)="VIEW_SPECIALITY_CLASS"
ArrayList(k,2)="CLASS_ID"
ArrayList(k,3)="CLASS_NAME"
ArrayList(k,4)=class_id
ArrayList(k,5)=""
Get_ListJavaMenu ArrayList,k,FormName,""
%></tr></table></td></tr>
<tr><td colspan=2><table cellspacing="4" cellpadding="0"><tr><td>表格审核状态</td><td><select name="In_TASK_PROGRESS"><option value="-1">所有</option><%
GetMenuListPubTerm "CODE_THESIS_REVIEW_STATUS","STATUS_ID1","STATUS_NAME",query_task_progress,"AND STATUS_ID1 IS NOT NULL"
%></select></td><td>论文审核状态</td><td><select name="In_REVIEW_STATUS"><option value="-1">所有</option><%
GetMenuListPubTerm "CODE_THESIS_REVIEW_STATUS","STATUS_ID2","STATUS_NAME",query_review_status,"AND STATUS_ID2 IS NOT NULL"
%></select></td></tr></table></td></tr><tr><td colspan=2>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_STU_NAME">学生姓名</option>
<option value="s_STU_NO">学号</option>
<option value="s_THESIS_SUBJECT">论文题目</option>
<option value="s_TUTOR_NAME">导师姓名</option>
<option value="ms_EXPERT_NAME1|EXPERT_NAME2">专家姓名</option>
</select>
<select name="operator">
<script>ReloadOperator()</script>
</select>
<input type="text" name="filter" size="10" onkeypress="checkKey()">
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>">
<input type="submit" value="查找" onclick="genFilter()">
<input type="submit" value="在结果中查找" onclick="genFinalFilter()"><%
If Len(PubTerm) Then %>
&nbsp;每页
<select name="pageSize" onchange="if(Chk_Select())submitForm(this.form)">
<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
</select>
条
&nbsp;
转到
<select name="pageNo" onchange="if(Chk_Select())submitForm(this.form)">
<%
For i=1 to rs.PageCount
    Response.write "<option value="&i
    If rs.AbsolutePage=i Then Response.write " selected"
    Response.write ">"&i&"</option>"
Next
%>
</select>
页
&nbsp;
共<%=rs.RecordCount%>条<%
End If %>
<input type="button" value="显示全部" onclick="showAllRecords(this.form)">
&nbsp;全选<input type="checkbox" onclick="checkAll()" id="chk" /></td></tr>
<tr><td colspan=2><input type="button" value="导入新增论文信息" onclick="submitForm($('#fmThesisList'),'importNewThesis.asp')" />
<input type="button" value="导入论文查重信息" onclick="submitForm($('#fmThesisList'),'importThesisDetInfo.asp')" />
<input type="button" value="导入答辩安排信息" onclick="submitForm($('#fmThesisList'),'importDefencePlan.asp')" />
<input type="button" value="导入答辩委员会修改意见" onclick="submitForm($('#fmThesisList'),'importDefenceEval.asp')" />
<input type="button" value="导入学院学位评定分委员会修改意见" onclick="submitForm($('#fmThesisList'),'importInstructEval.asp')" /></td></tr>
<tr><td colspan=2>评阅结果&nbsp;<select name="selreviewfilestat"><%
			For i=0 To UBound(arrReviewFileStat)
%><option value="<%=i%>"><%=arrReviewFileStat(i)%></option><%
			Next %></select><input type="button" value="设置" onclick="batchUpdateThesis($('#fmThesisList'))" />&emsp;
<input type="button" value="导入专家匹配结果" onclick="submitForm($('#fmThesisList'),'importMatchResult.asp')" />
<input type="button" value="批量通知专家评阅" onclick="submitForm($('#fmThesisList'),'notifyExpert.asp')" />
<input type="button" value="批量下载表格/论文" onclick="batchFetchFile($('#fmThesisList'))" />
<input type="button" name="btnexport" value="导出到Excel文件" />
<input type="button" value="删 除" onclick="if(confirm('是否删除这'+countClk()+'条记录？'))submitForm($('#fmThesisList'),'delThesis.asp')" /></td></tr></form></table>
<form id="fmThesisList" method="post">
<input type="hidden" name="In_PERIOD_ID2" value="<%=period_id%>">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>">
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>">
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>">
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>">
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>">
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>">
<input type="hidden" name="pageSize2" value=<%=pageSize%>>
<input type="hidden" name="pageNo2" value=<%=pageNo%>>
<input type="hidden" name="reviewfilestat" value="0">
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
  <tr bgcolor="gainsboro" height="25">
    <td align=center>论文题目</td>
    <td width="80" align=center>姓名</td>
    <td width="90" align=center>学号</td>
    <td width="120" align=center>专业</td>
    <td width="50" align=center>学位类别</td>
    <td width="60" align=center>导师</td><%
    If Not bModified Then %>
		<td width="80" align=center>送审结果1</td>
		<td width="80" align=center>送审结果2</td><%
		End If %>
		<td width="80" align=center>处理意见</td>
		<td width="110" align=center>状态</td>
    <td width="30" align=center>选择</td>
    <td width="100" align=center>操作</td>
  </tr><%
If Len(PubTerm) Then
  Dim review_result
  For i=1 to rs.PageSize
    If rs.EOF Then Exit For
    If Not IsNull(rs("REVIEW_RESULT")) Then
    	review_result=Split(rs("REVIEW_RESULT"),",")
    Else
    	ReDim review_result(2)
    End If
    substat=vbNullString
		If rs("TASK_PROGRESS")>=tpTbl4Uploaded Then
    	stat=rs("STAT_TEXT1")&"，"&rs("STAT_TEXT2")
		ElseIf rs("REVIEW_STATUS")=0 Then
    	stat=rs("STAT_TEXT1")
    Else
    	stat=rs("STAT_TEXT2")
    	If rs("REVIEW_STATUS")>=rsReviewed And rs("REVIEW_FILE_STATUS")<>3 Then
    		substat="评阅结果["&arrReviewFileStat(rs("REVIEW_FILE_STATUS"))&"]"
    	End If
  	End If
  	If rs("UNHANDLED") Then
  		cssclass="thesisstat_unhandled"
  	Else
  		cssclass="thesisstat"
  	End If
  %><tr bgcolor="ghostwhite">
    <td align=center><a href="#" onclick="tabmgr.newTab('/ThesisReview/admin/thesisDetail.asp?tid=<%=rs("ID")%>');return false;"><%=HtmlEncode(rs("THESIS_SUBJECT"))%></a></td>
    <td align=center><%=HtmlEncode(rs("STU_NAME"))%></td>
    <td align=center><%=rs("STU_NO")%></td>
    <td align=center><%=HtmlEncode(rs("SPECIALITY_NAME"))%></td>
    <td align=center><%=rs("TEACHTYPE_NAME")%></td>
    <td align=center><%=HtmlEncode(rs("TUTOR_NAME"))%></td><%
    If Not bModified Then %>
		<td align=center><%=getReviewResult(review_result(0))%></td>
    <td align=center><%=getReviewResult(review_result(1))%></td><%
		End If %>
    <td align=center><%=getFinalResult(review_result(2))%></td>
    <td align=center><a href="#" onclick="tabmgr.newTab('/ThesisReview/admin/thesisDetail.asp?tid=<%=rs("ID")%>');return false;"><span class="<%=cssclass%>"><%=stat%></span></a><%
    If Len(substat) Then
    %><br/><span class="thesissubstat"><%=substat%></span><%
    End If %></td>
    <td align=center><input type="checkbox" name="sel" value="<%=rs("ID")%>" /></td>
    <td align=center><%
  	If rs("REVIEW_STATUS")=rsAgreeReview Then
  		If IsNull(rs("REVIEWER1")) And IsNull(rs("REVIEWER2")) Then
%><input type="button" value="匹配专家" onclick="chooseExpert(this.form,<%=rs("ID")%>)" /><%
  		Else
%><input type="button" value="通知专家评阅" onclick="notifyExpert(this.form,<%=rs("ID")%>)" /><%
  		End If
  	End If
  	%></td></tr><%
  	rs.MoveNext()
  Next
End If
%></table></form></center></body>
<script type="text/javascript">
	document.all.btnexport.onclick=function() {
		this.value="正在导出，请稍候……";
		this.disabled=true;
		submitForm($('#fmThesisList'),"exportReviewStats.asp");
	}
	document.all.btnexport.disabled=false;
</script></html><%
  CloseRs rs
  CloseConn conn
%>