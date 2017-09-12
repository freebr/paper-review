<%Response.Charset="utf-8"
Response.Expires=-1
Response.Buffer=True
%>
<!--#include file="inc/db.asp"-->
<!--#include file="teacher/common.asp"-->
<%
	sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_INFO WHERE REVIEW_APP IS NULL AND REVIEW_STATUS>="&rsAgreeReview
	GetRecordSet conn,rs,sql,result
	n=0
	Do While Not rs.EOF
		Response.Write (n+1)&":正在生成 "&rs("STU_NAME")&" 的送审申请表……"
		Response.Flush()
		' 生成送审申请表
		Dim rag
		Randomize
		review_time=rs("SUBMIT_REVIEW_TIME")
		eval_text=rs("REVIEW_APP_EVAL")
		If IsNull(eval_text) Then eval_text=""
		filename=toDateTime(review_time,1)&Int(Timer)&Int(Rnd()*999)&".doc"
		filepath=Server.MapPath("/ThesisReview/teacher/export")&"\"&filename
		Set rag=New ReviewAppGen
		rag.Author=rs("STU_NAME")
		rag.StuNo=rs("STU_NO")
		rag.TutorInfo=rs("TUTOR_NAME")&" "&getProDutyNameOf(rs("TUTOR_ID"))
		rag.Spec=rs("SPECIALITY_NAME")
		rag.Date=toDateTime(review_time,1)
		rag.Subject=rs("THESIS_SUBJECT")
		rag.EvalText=eval_text
		rag.ReproductRatio=rs("REPRODUCTION_RATIO")
		bError=rag.generateApp(filepath)=0
		Set rag=Nothing
		rs("REVIEW_APP")=filename
		rs.Update()
		rs.MoveNext()
		Response.Write filename&" 成功！<br/>"
		Response.Flush()
		n=n+1
	Loop
	Response.Write "更新成功，共更新了 "&n&" 条记录。"
	CloseRs rs
	CloseConn conn
%>