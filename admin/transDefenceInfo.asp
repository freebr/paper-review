<!--#include file="../inc/db.asp"--><%
Dim count:count=0

'Response.Charset="utf-8"
sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE DEFENCE_RESULT <>0 OR DEFENCE_MODIFY_EVAL IS NOT NULL"
GetRecordSet conn,rs,sql,result

Response.Write "正在转移答辩结果……"
Do While Not rs.EOF
	Response.Write "<br>正在转移学生["&rs("STU_NAME")&"]的论文["&rs("THESIS_SUBJECT")&"]的答辩结果……"
	sql2="IF NOT EXISTS(SELECT THESIS_ID FROM TEST_THESIS_DEFENCE_INFO WHERE THESIS_ID="&rs("ID")&") INSERT INTO TEST_THESIS_DEFENCE_INFO (THESIS_ID,DEFENCE_RESULT) VALUES("&_
				rs("ID")&","&rs("DEFENCE_RESULT")&")"
	conn.Execute sql2
	
	sql2="SELECT * FROM TEST_THESIS_DEFENCE_INFO WHERE THESIS_ID="&rs("ID")
	GetRecordSet conn,rs2,sql2,result
	rs2("DEFENCE_EVAL")=rs("DEFENCE_MODIFY_EVAL")
	rs2.Update()
	CloseRs rs2
	
	Response.Write "完成！"
	Response.Flush()
	count=count+1
	rs.MoveNext()
Loop
Response.Write "共转移了 "&count&" 条答辩结果。"
CloseRs rs
CloseConn conn
%>