<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
uid_type=Request("sel")
finalFilter=Request.Form("finalFilter2")
pageSize=Request.Form("pageSize2")
pageNo=Request.Form("pageNo2")
If Len(uid_type)=0 Then
	bError=True
	errdesc="您未选择通知对象！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End()
End If

Dim arr,arr2,user_type,user_id
Dim conn,rs,sql,result
Dim mail_id
Dim fieldval,logtxt,bAllowSms,bMailSuccess,bMobSuccess
Dim numTutorNotify,numExpertNotify,numFailed,errMsg
mail_id=getThesisReviewSystemMailIdByType(Now)
numTutorNotify=0
numExpertNotify=0
numFailed=0
bAllowSms=True

arr=Split(uid_type,",")
Connect conn
For i=0 To UBound(arr)
	arr2=Split(arr(i),".")
	user_type=arr2(0)
	user_id=arr2(1)
	sql="SELECT * FROM VIEW_TEST_THESIS_REVIEW_NOTIFY_INFO WHERE USER_ID="&user_id
	GetRecordSetNoLock conn,rs,sql,result
	If Not rs.EOF Then
		Select Case user_type
		Case 1	' 导师
			
			tutorname=rs("USER_NAME")
			tutormail=rs("USER_EMAIL")
			notify_desc="您有 "
	    If rs("AUDIT_COUNT")>0 Then
	    	notify_desc=notify_desc&rs("AUDIT_COUNT")&" 份表格/论文待审核："
	  	End If
	    If rs("REQUEST_REVIEW_COUNT")>0 Then
	    	notify_desc=notify_desc&rs("REQUEST_REVIEW_COUNT")&" 份论文已由教务员匹配专家并送审；"
	  	End If
	    If rs("INFO_IMPORTED_COUNT")>0 Then
	    	notify_desc=notify_desc&rs("INFO_IMPORTED_COUNT")&" 份论文已给出答辩安排（意见）；"
	  	End If
	  	notify_desc=Left(notify_desc,Len(notify_desc)-1)
  		fieldval=Array(tutorname,tutormail,notify_desc)
			logtxt="行政人员["&Session("name")&"]通知导师待办事项，"
			bMailSuccess=sendAnnouncementEmail(mail_id(11),tutormail,fieldval)
			logtxt=logtxt&"发送邮件给导师["&tutorname&":"&tutormail&"]"
			If bMailSuccess Then
				numTutorNotify=numTutorNotify+1
				logtxt=logtxt&"成功。"
			Else
				numFailed=numFailed+1
				errMsg=errMsg&"\r\n向["&tutorname&"]发送通知失败，邮箱["&tutormail&"]。"
				logtxt=logtxt&"失败。"
			End If
			
		Case 2	' 评阅专家
			
			expertid=rs("USER_ID")
			expertname=rs("USER_NAME")
			expertmail=rs("USER_EMAIL")
			expertmob=rs("USER_MOBILE")
			sql="SELECT A.TEACHERNO,COUNT(ID) AS REVIEW_COUNT,MIN(THESIS_SUBJECT) AS THESIS_SUBJECT FROM VIEW_TEST_THESIS_REVIEW_INFO "&_
					"LEFT JOIN VIEW_TEACHER_INFO A ON A.TEACHERID="&expertid&" WHERE "&expertid&" IN (REVIEWER1,REVIEWER2) GROUP BY TEACHERNO"
			GetRecordSetNoLock conn,rs2,sql,result
			review_count=rs2("REVIEW_COUNT")
			If review_count=1 Then
				subject="《"&rs2("THESIS_SUBJECT")&"》"
			Else
				subject="《"&rs2("THESIS_SUBJECT")&"》等"&review_count&"篇论文"
			End If
			postscript="您的登录名为&nbsp;<b>"&rs2("TEACHERNO")&"</b>，初始密码为&nbsp:<b>123456</b>，登录后请务必修改您的密码"
			' 发送通知短信和邮件
			fieldval=Array(subject,expertname,expertmob,expertmail,postscript)
			logtxt="行政人员["&Session("name")&"]通知专家评阅论文，"
			ret=-1
			If bAllowSms Then
				ret=sendSMS(mail_id(4),expertmob,fieldval)
				logtxt=logtxt&"发送短信给评阅专家["&expertname&":"&expertmob&"]"
				bMobSuccess=ret=0
				If bMobSuccess Then
					logtxt=logtxt&"成功。"
				Else
					logtxt=logtxt&"失败("&ret&")。"
				End If
			Else
				bMobSuccess=False
			End If
			bMailSuccess=sendAnnouncementEmail(mail_id(3),expertmail,fieldval)
			logtxt=logtxt&"发送邮件给评阅专家["&expertname&":"&expertmail&"]"
			If bMailSuccess Then
				logtxt=logtxt&"成功。"
			Else
				logtxt=logtxt&"失败。"
			End If
			If bMobSuccess Or bMailSuccess Then
				numExpertNotify=numExpertNotify+1
			Else
				numFailed=numFailed+1
				errMsg=errMsg&"\r\n向["&expertname&"]发送通知失败，手机["&expertmob&"]，邮箱["&expertmail&"]。"
			End If
			CloseRs rs2
			
		End Select
		
		' 更新通知情况
		sql_updnotify=sql_updnotify&"UPDATE TEST_THESIS_REVIEW_NOTIFY_INFO SET LAST_NOTIFY_TIME="&toSqlString(Now)&_
																",LAST_NOTIFY_MAIL_RESULT="&Abs(Int(bMailSuccess))&",LAST_NOTIFY_MOB_RESULT="&Abs(Int(bMobSuccess))&_
																" WHERE USER_ID="&user_id&";"
	End If
Next

If Len(sql_updnotify) Then conn.Execute sql_updnotify
CloseConn conn
WriteLog logtxt
notifyText="操作完成，共通知 "
If numTutorNotify>0 Then
	notifyText=notifyText&numTutorNotify&" 名导师，"
End If
If numExpertNotify>0 Then
	notifyText=notifyText&numExpertNotify&" 名专家，"
End If
If numFailed>0 Then
	notifyText=notifyText&"其中 "&numFailed&" 名发送通知失败，原因如下："&errMsg
Else
	notifyText=notifyText&"全部通知成功。"
End If
%><form id="ret" action="notifyList.asp" method="post">
<input type="hidden" name="finalFilter" value="<%=toPlainString(finalFilter)%>" />
<input type="hidden" name="pageSize" value="<%=pageSize%>" />
<input type="hidden" name="pageNo" value="<%=pageNo%>" /></form>
<script type="text/javascript">
	alert("<%=notifyText%>");
	document.all.ret.submit();
</script>