<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
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

Dim mail_id
Dim numNotify,numSuccess,errMsg
mail_id=getThesisReviewSystemMailIdByType(Now)
numNotify=0
numSuccess=0
bAllowSms=True
logtxt="行政人员["&Session("name")&"]通知专家评阅论文，"
Connect conn
sql="DECLARE @tmptbl TABLE(EXPERT_ID int,THESIS_SUBJECT nvarchar(200),REVNUM int);"&_
	"INSERT INTO @tmptbl SELECT REVIEWER1,MIN(THESIS_SUBJECT),COUNT(ID) FROM ViewThesisInfo WHERE ID IN ("&thesisID&") AND REVIEWER_EVAL1 IS NULL GROUP BY REVIEWER1 "&_
	"UNION ALL SELECT REVIEWER2,MIN(THESIS_SUBJECT),COUNT(ID) FROM ViewThesisInfo WHERE ID IN ("&thesisID&") AND REVIEWER_EVAL2 IS NULL GROUP BY REVIEWER2;"&_
	"SELECT EXPERT_ID,EXPERT_NAME,TEACHERNO,MOBILE,EMAIL,MIN(THESIS_SUBJECT) AS THESIS_SUBJECT,SUM(REVNUM) AS REVIEW_NUM FROM @tmptbl LEFT JOIN ViewExpertInfo ON EXPERT_ID=TEACHER_ID GROUP BY EXPERT_ID,EXPERT_NAME,TEACHERNO,MOBILE,EMAIL;"
Set rs=conn.Execute(sql)
Set rs=rs.NextRecordSet
Do While Not rs.EOF
	review_num=rs("REVIEW_NUM")
	If review_num=1 Then
		subject="《"&rs("THESIS_SUBJECT")&"》"
	Else
		subject="《"&rs("THESIS_SUBJECT")&"》等"&review_num&"篇论文"
	End If
	expertname=rs("EXPERT_NAME")
	expertmob=rs("MOBILE")
	expertmail=rs("EMAIL")
	postscript="您的登录名为&nbsp;<b>"&rs("TEACHERNO")&"</b>，初始密码为&nbsp:<b>123456</b>，登录后请务必修改您的密码"
	' 发送通知短信和邮件
	fieldval=Array(subject,expertname,expertmob,expertmail,postscript)
	ret=-1
	If bAllowSms Then
		ret=sendSMS(mail_id(3),expertmob,fieldval)
		logtxt=logtxt&"发送短信给评阅专家["&expertname&":"&expertmob&"]"
		If ret=0 Then
			logtxt=logtxt&"成功。"
		Else
			logtxt=logtxt&"失败("&ret&")。"
		End If
	End If
	bSuccess=sendAnnouncementEmail(mail_id(3),expertmail,fieldval)
	logtxt=logtxt&"发送邮件给评阅专家["&expertname&":"&expertmail&"]"
	If bSuccess Then
		logtxt=logtxt&"成功。"
	Else
		logtxt=logtxt&"失败。"
	End If
	If ret=0 Or bSuccess Then
		numSuccess=numSuccess+1
	Else
		errMsg=errMsg&"\r\n向["&expertname&"]发送通知失败，登记信息：手机["&expertmob&"]，邮箱["&expertmail&"]。"
	End If
	numNotify=numNotify+1
	rs.MoveNext()
Loop
CloseRs rs
CloseConn conn

writeLog logtxt
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