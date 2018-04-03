<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"--><%
If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
session("Debug")=true
Dim newTutorCount,newExpertCount:newTutorCount=0:newExpertCount=0
Dim bError,errMsg:bError=False
Connect conn
sql="DELETE FROM TEST_THESIS_REVIEW_NOTIFY_INFO"
conn.Execute sql

' 导入不在通知列表中的待通知导师记录
sql="SELECT DISTINCT TUTOR_ID FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE (TASK_PROGRESS IN (1,4,7,10) OR REVIEW_STATUS IN (1,5,8,11,14,15)) "&_
		"AND TUTOR_ID IS NOT NULL AND TUTOR_ID NOT IN (SELECT USER_ID FROM TEST_THESIS_REVIEW_NOTIFY_INFO WHERE USER_TYPE=1)"
GetRecordSetNoLock conn,rs,sql,result
sql="INSERT INTO TEST_THESIS_REVIEW_NOTIFY_INFO (USER_ID,USER_TYPE) VALUES"
Do While Not rs.EOF
	If newTutorCount>0 Then sql=sql&","
	sql=sql&"("&rs(0)&",1)"
	newTutorCount=newTutorCount+1
	rs.MoveNext()
Loop
conn.Execute sql

' 导入不在通知列表中的待通知专家记录
sql="SELECT DISTINCT REVIEWER1 FROM TEST_THESIS_REVIEW_INFO WHERE REVIEWER1 IS NOT NULL AND REVIEWER_EVAL1 IS NULL "&_
		"UNION SELECT DISTINCT REVIEWER2 FROM TEST_THESIS_REVIEW_INFO WHERE REVIEWER2 IS NOT NULL AND REVIEWER_EVAL2 IS NULL"
GetRecordSetNoLock conn,rs,sql,result
sql="INSERT INTO TEST_THESIS_REVIEW_NOTIFY_INFO (USER_ID,USER_TYPE) VALUES"
Do While Not rs.EOF
	If newExpertCount>0 Then sql=sql&","
	sql=sql&"("&rs(0)&",2)"
	newExpertCount=newExpertCount+1
	rs.MoveNext()
Loop
conn.Execute sql

CloseRs rs
CloseConn conn

%><script type="text/javascript"><%
	If bError Then %>
	alert("导入时出错，其余<%=newTutorCount%>名导师和<%=newExpertCount%>名专家已导入成功。出错原因为：\n<%=toJsString(errMsg)%>");
<%Else %>
	alert("操作成功，已导入<%=newTutorCount%>名导师和<%=newExpertCount%>名专家。");
<%End If
%>location.href="notifyList.asp";
</script>