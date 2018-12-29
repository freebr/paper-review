<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
usertype=Request.QueryString("user")
opr=Request.QueryString("opr")
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(thesisID)=0 Or Not IsNumeric(thesisID) Or Len(usertype)=0 Or Not IsNumeric(usertype) Or Len(opr)=0 Or Not IsNumeric(opr) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">参数无效。</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim conn,rs,sql,sqlDetect,result
Connect conn
sql="SELECT * FROM TEST_THESIS_REVIEW_INFO WHERE ID="&thesisID
GetRecordSet conn,rs,sql,result
If rs.EOF Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">数据库没有该论文记录！</font><br/><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
  CloseRs rs
  CloseConn conn
	Response.End()
End If

Dim detect_count
sql="SELECT DETECT_COUNT FROM ViewThesisInfo WHERE ID="&thesisID
GetRecordSet conn,rsDetect,sql,result
detect_count=rsDetect(0).Value
CloseRs rsDetect

Select Case usertype
Case 0	' 撤销学生上传操作
	Select Case opr
	Case 0
		rs("TABLE_FILE1").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TBL_THESIS_FILE1").Value) Then
			rs("TASK_PROGRESS").Value=0
		Else
			rs("TASK_PROGRESS").Value=tpTbl1Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 1
		rs("TBL_THESIS_FILE1").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TABLE_FILE1").Value) Then
			rs("TASK_PROGRESS").Value=0
		Else
			rs("TASK_PROGRESS").Value=tpTbl1Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 2
		rs("TABLE_FILE2").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TBL_THESIS_FILE2").Value) Then
			rs("TASK_PROGRESS").Value=tpTbl1Passed
		Else
			rs("TASK_PROGRESS").Value=tpTbl2Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 3
		rs("TBL_THESIS_FILE2").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TABLE_FILE2").Value) Then
			rs("TASK_PROGRESS").Value=tpTbl1Passed
		Else
			rs("TASK_PROGRESS").Value=tpTbl2Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 4
		rs("TABLE_FILE3").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TBL_THESIS_FILE3").Value) Then
			rs("TASK_PROGRESS").Value=tpTbl2Passed
		Else
			rs("TASK_PROGRESS").Value=tpTbl3Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 5
		rs("TBL_THESIS_FILE3").Value=Null
		'rs("TASK_EVAL").Value=Null
		If IsNull(rs("TABLE_FILE3").Value) Then
			rs("TASK_PROGRESS").Value=tpTbl2Passed
		Else
			rs("TASK_PROGRESS").Value=tpTbl3Uploaded
		End If
		rs("REVIEW_STATUS").Value=0
	Case 6
		rs("THESIS_FILE").Value=Null
		'rs("DETECT_APP_EVAL").Value=Null
		rs("TASK_PROGRESS").Value=tpTbl3Passed
		rs("REVIEW_STATUS").Value=0
	Case 7
		rs("THESIS_FILE2").Value=Null
		'rs("REVIEW_APP_EVAL").Value=Null
		rs("REVIEW_STATUS").Value=rsAgreeDetect
	Case 8
		rs("THESIS_FILE3").Value=Null
		'rs("TUTOR_MODIFY_EVAL").Value=Null
		rs("REVIEW_STATUS").Value=rsReviewEval
	Case 9
		rs("THESIS_FILE4").Value=Null
		rs("REVIEW_STATUS").Value=rsInstructEval
	Case 10
		rs("TABLE_FILE4").Value=Null
		'rs("TASK_EVAL").Value=Null
		rs("TASK_PROGRESS").Value=tpTbl3Passed
	End Select
Case 1	' 撤销专家评阅操作

	Dim review_result(2),reviewer_master_level(1),review_file(1),review_time(1),review_level(1)
	If Not IsNull(rs("REVIEW_RESULT").Value) Then
		arr=Split(rs("REVIEW_RESULT").Value,",")
		For i=0 To UBound(arr)
			review_result(i)=Int(arr(i))
		Next
	End If
	If IsNull(rs("REVIEWER_EVAL_TIME").Value) Then
		For i=0 To 1
			reviewer_master_level(i)=0
			review_level(i)=0
		Next
	Else
		arr2=Split(rs("REVIEWER_MASTER_LEVEL").Value,",")
		arr3=Split(rs("REVIEW_FILE").Value,",")
		arr4=Split(rs("REVIEWER_EVAL_TIME").Value,",")
		arr5=Split(rs("REVIEW_LEVEL").Value,",")
		For i=0 To 1
			reviewer_master_level(i)=Int(arr2(i))
			review_file(i)=arr3(i)
			review_time(i)=arr4(i)
			review_level(i)=Int(arr5(i))
		Next
	End If
	
	review_result(opr)="5"
	review_level(opr)="0"
	reviewer_master_level(opr)="0"
	review_file(opr)=""
	review_time(opr)=""
	finalresult="6"
	review_result(2)=finalresult
	
	' 更新记录
	rs("REVIEW_RESULT").Value=join(review_result,",")
	rs("REVIEW_LEVEL").Value=join(review_level,",")
	rs("REVIEWER_MASTER_LEVEL").Value=join(reviewer_master_level,",")
	rs("REVIEWER_EVAL"&(opr+1)).Value=eval_text
	rs("REVIEW_FILE").Value=join(review_file,",")
	rs("REVIEWER_EVAL_TIME").Value=join(review_time,",")
	'rs("TUTOR_REVIEW_EVAL").Value=Null
	rs("REVIEW_STATUS").Value=rsMatchExpert
	
Case 2	' 撤销导师审核操作
	Select Case opr
	Case 0
		rs("TASK_EVAL").Value=Null
		Select Case rs("TASK_PROGRESS").Value
		Case tpTbl1Unpassed,tpTbl1Passed
			rs("TASK_PROGRESS").Value=tpTbl1Uploaded
			rs("REVIEW_STATUS").Value=0
		Case tpTbl2Unpassed,tpTbl2Passed
			rs("TASK_PROGRESS").Value=tpTbl2Uploaded
			rs("REVIEW_STATUS").Value=0
		Case tpTbl3Unpassed,tpTbl3Passed
			rs("TASK_PROGRESS").Value=tpTbl3Uploaded
			rs("REVIEW_STATUS").Value=0
		Case tpTbl4Unpassed,tpTbl4Passed
			rs("TASK_PROGRESS").Value=tpTbl4Uploaded
		End Select
	Case 1
		rs("DETECT_APP_EVAL").Value=Null
		rs("REVIEW_STATUS").Value=rsDetectThesisUploaded
	Case 2
		rs("REVIEW_APP").Value=Null
		rs("REVIEW_APP_EVAL").Value=Null
		rs("SUBMIT_REVIEW_TIME").Value=Null
		If detect_count>1 Then	' 检测次数超过一次，回退到二次检测通过状态
			rs("REVIEW_STATUS").Value=rsRedetectPassed
		Else	' 一次检测通过，回退到待同意检测状态
			rs("DETECT_APP_EVAL").Value=Null
			rs("REVIEW_STATUS").Value=rsDetectThesisUploaded
		End If
	Case 3
		rs("TUTOR_MODIFY_EVAL").Value=Null
		rs("REVIEW_STATUS").Value=rsModifyThesisUploaded
	End Select
Case 3	' 撤销教务员操作
	Select Case opr
	Case 0	' 撤销送检操作
		sqlDetect="EXEC spDeleteDetectResult "&thesisID&";"
		If Not IsNull(rs("THESIS_FILE").Value) Then
			sqlDetect=sqlDetect&"EXEC spAddDetectResult "&thesisID&","&rs("THESIS_FILE").Value&",NULL,NULL,NULL;"
		End If
		rs("REVIEW_STATUS").Value=rsAgreeDetect
	Case 1	' 撤销匹配专家操作
		rs("REVIEW_STATUS").Value=rsAgreeReview
	Case 2	' 撤销导入答辩安排操作
		sql="DELETE FROM TEST_THESIS_DEFENCE_INFO WHERE THESIS_ID="&thesisID
		conn.Execute sql
	Case 3	' 撤销导入答辩委员会修改意见操作
		sql="UPDATE TEST_THESIS_DEFENCE_INFO SET DEFENCE_EVAL=NULL WHERE THESIS_ID="&thesisID
		conn.Execute sql
		rs("DEFENCE_MODIFY_EVAL").Value=Null	' 旧字段
		rs("REVIEW_STATUS").Value=rsModifyPassed
	Case 4	' 撤销导入教指会分会修改意见操作
		rs("INSTRUCT_MODIFY_EVAL").Value=Null
		rs("REVIEW_STATUS").Value=rsDefenceEval
	End Select
End Select

If rs("REVIEW_STATUS").Value<rsMatchExpert Then
	rs("REVIEWER1").Value=Null
	rs("REVIEWER2").Value=Null
	rs("REVIEW_RESULT").Value="5,5,6"
	rs("REVIEW_LEVEL").Value="0,0"
	rs("REVIEWER_EVAL1").Value=Null
	rs("REVIEWER_EVAL2").Value=Null
	rs("REVIEWER_MASTER_LEVEL").Value=Null
	rs("REVIEW_FILE").Value=Null
	rs("REVIEW_FILE_STATUS").Value=0
	rs("REVIEWER_EVAL_TIME").Value=Null
End If
rs.Update()
If Len(sqlDetect) Then
	conn.Execute sqlDetect
End If
CloseRs rs
CloseConn conn
%><form id="ret" action="thesisDetail.asp?tid=<%=thesisID%>" method="post">
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成。");
	document.all.ret.submit();
</script>