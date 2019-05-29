<%Response.Expires=-1%>
<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
thesisID=Request.QueryString("tid")
If IsEmpty(thesisId) Then
	thesisID=Request.Form("sel")
End If
teachtype_id=Request.Form("In_TEACHTYPE_ID2")
class_id=Request.Form("In_CLASS_ID2")
enter_year=Request.Form("In_ENTER_YEAR2")
query_task_progress=Request.Form("In_TASK_PROGRESS2")
query_review_status=Request.Form("In_REVIEW_STATUS2")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(thesisID)=0 Then thesisID=Request.Form("thesisID")
If Len(thesisID)=0 Then
	bError=True
	errdesc="您未选择论文！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim numNotify:numNotify=0
Dim numSuccess:numSuccess=0
Dim send_sms:send_sms=True
Dim dict:Set dict=CreateDictionary()
Dim activity_id,stu_type,is_mail_sent,is_sms_sent
Dim errMsg
Connect conn
sql="DECLARE @tmptbl TABLE(EXPERT_ID int,ActivityId int,StuType int,THESIS_SUBJECT nvarchar(200),REVNUM int);"&_
	"INSERT INTO @tmptbl SELECT REVIEWER1,MAX(ActivityId),MIN(TEACHTYPE_ID),MIN(THESIS_SUBJECT),COUNT(ID) FROM ViewThesisInfo WHERE ID IN ("&thesisID&") AND REVIEWER_EVAL1 IS NULL GROUP BY REVIEWER1 "&_
	"UNION ALL SELECT REVIEWER2,MAX(ActivityId),MIN(TEACHTYPE_ID),MIN(THESIS_SUBJECT),COUNT(ID) FROM ViewThesisInfo WHERE ID IN ("&thesisID&") AND REVIEWER_EVAL2 IS NULL GROUP BY REVIEWER2;"&_
	"SELECT EXPERT_ID,EXPERT_NAME,TEACHERNO,MOBILE,EMAIL,MAX(ActivityId) ActivityId,MIN(StuType) StuType,MIN(THESIS_SUBJECT) THESIS_SUBJECT,SUM(REVNUM) REVIEW_NUM FROM @tmptbl LEFT JOIN ViewExpertInfo ON EXPERT_ID=TEACHER_ID GROUP BY EXPERT_ID,EXPERT_NAME,TEACHERNO,MOBILE,EMAIL;"
Set rs=conn.Execute(sql)
Set rs=rs.NextRecordSet()
Do While Not rs.EOF
	activity_id=rs("ActivityId")
	stu_type=rs("StuType")
	review_num=rs("REVIEW_NUM")
	If review_num=1 Then
		dict("subject")=Format("《{0}》",rs("THESIS_SUBJECT"))
	Else
		dict("subject")=Format("《{0}》等{1}篇论文",rs("THESIS_SUBJECT"),review_num)
	End If
	dict("expertname")=rs("EXPERT_NAME")
	dict("expertmob")=rs("MOBILE")
	dict("expertmail")=rs("EMAIL")
	dict("postscript")=Format("您的登录名为 <b>{0}</b>，初始密码为 <b>123456</b>，登录后请务必修改您的密码",rs("TEACHERNO"))
	' 发送通知邮件
	is_mail_sent=sendNotifyMail(activity_id,stu_type,"lwdpytzyj",dict("expertmail"),dict)
	writeNotificationEventLog usertypeAdmin,Session("name"),"通知专家评阅论文",usertypeExpert,_
		dict("expertname"),dict("expertmail"),notifytypeMail,is_mail_sent
	If send_sms Then
		' 发送通知短信
		is_sms_sent=sendNotifySms(activity_id,stu_type,"lwdpytzdx",dict("expertmob"),dict)
		writeNotificationEventLog usertypeAdmin,Session("name"),"通知专家评阅论文",usertypeExpert,_
			dict("expertname"),dict("expertmob"),notifytypeSms,is_sms_sent
	Else
		is_sms_sent=False
	End If
	If is_mail_sent Or is_sms_sent Then
		numSuccess=numSuccess+1
	Else
		errMsg=errMsg&Format("\r\n向[{0}]发送通知失败，手机：{1}，邮箱：{2}。",_
			dict("expertname"),dict("expertmob"),dict("expertmail"))
	End If
	numNotify=numNotify+1
	rs.MoveNext()
Loop
CloseRs rs
CloseConn conn

If InStr(thesisID,",") Then
	returl="thesisList.asp"
Else
	returl="thesisDetail.asp?tid="&thesisID
End If
%><form id="ret" action="<%=returl%>" method="post">
<input type="hidden" name="In_TEACHTYPE_ID" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" />
<input type="hidden" name="In_TEACHTYPE_ID2" value="<%=teachtype_id%>" />
<input type="hidden" name="In_CLASS_ID2" value="<%=class_id%>" />
<input type="hidden" name="In_ENTER_YEAR2" value="<%=enter_year%>" />
<input type="hidden" name="In_TASK_PROGRESS2" value="<%=query_task_progress%>" />
<input type="hidden" name="In_REVIEW_STATUS2" value="<%=query_review_status%>" />
<input type="hidden" name="finalFilter2" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize2" value="<%=pageSize%>" />
<input type="hidden" name="pageNo2" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("操作完成，共通知 <%=numNotify%> 名专家，其中 <%=numSuccess%> 名发送通知成功。<%=errMsg%>");
	document.all.ret.submit();
</script>