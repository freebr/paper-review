<%
Function newEmailTemplate(template_name,mailsubject,mailcontent,fieldlist)
	' 新建邮件模板并返回模板编号
	Dim sql,rs,conn,result,updateTime
	updateTime="'"&Now&"'"
	sql="INSERT INTO EMAIL_TEMPLATE (TemplateName,MailSubject,MailContent,FieldList,UpdateTime,Valid) "&_
			"VALUES ("&toSqlString(template_name)&","&toSqlString(mailsubject)&","&toSqlString(mailcontent)&","&toSqlString(fieldlist)&","&updateTime&",1)"
	Connect conn
	conn.Execute sql
	sql="SELECT ID FROM EMAIL_TEMPLATE WHERE UpdateTime="&updateTime&" ORDER BY ID DESC"
	GetRecordSetNoLock conn,rs,sql,result
	If result>0 Then
		newEmailTemplate=rs(0)
	End If
	CloseRs rs
	CloseConn conn
End Function
Function updateEmailTemplate(id,template_name,mailsubject,mailcontent,fieldlist)
	' 修改指定编号的邮件模板
	If Not IsNumeric(id) Then Exit Function
	Dim sql,rs,conn,result,updateTime
	
	sql="SELECT ID FROM EMAIL_TEMPLATE WHERE ID="&id
	Connect conn
	GetRecordSetNoLock conn,rs,sql,result
	If result=0 Then
		id=newEmailTemplate(template_name,mailsubject,mailcontent,fieldlist)
	Else
		updateTime="'"&Now&"'"
		sql="UPDATE EMAIL_TEMPLATE SET TemplateName="&toSqlString(template_name)&","&_
																	"MailSubject="&toSqlString(mailsubject)&","&_
																	"MailContent="&toSqlString(mailcontent)&","&_
																	"FieldList="&toSqlString(fieldlist)&","&_
																	"UpdateTime="&updateTime&",Valid=1 WHERE ID="&id
		conn.Execute sql
	End If
	CloseRs rs
	CloseConn conn
	updateEmailTemplate=id
End Function
Function deleteEmailTemplate(id)
	' 删除指定编号的邮件模板
	If Not IsNumeric(id) Then Exit Function
	Dim sql,conn
	sql="DELETE FROM EMAIL_TEMPLATE WHERE ID="&id
	Connect conn
	conn.Execute sql
	CloseConn conn
	deleteEmailTemplate=1
End Function
Function getEmailTemplateContent(id)
	' 返回指定编号的邮件模板内容
	If IsNull(id) Or Len(id)=0 Or Not IsNumeric(id) Then Exit Function
	Dim sql,rs,conn
	sql="SELECT MailContent FROM EMAIL_TEMPLATE WHERE ID="&id
	Connect conn
	GetRecordSetNoLock conn,rs,sql,result
	If result>0 Then
		getEmailTemplateContent=rs(0)
	End If
	CloseRs rs
	CloseConn conn
End Function
Function sendAnnouncementEmail(mail_id,mailrcpt,arr_fieldval)
	If IsNull(mail_id) Then Exit Function
	Dim sql,rs,conn,result,i
	Dim arr_fieldname
	Dim jmail,mailsubject,mailcontent,mailbody
	Dim bSuccess
	
	'On Error Resume Next
	' 查询邮件模板标题、内容、变量等信息
	sql="SELECT * FROM EMAIL_TEMPLATE WHERE ID="&mail_id
	Connect conn
	GetRecordSetNoLock conn,rs,sql,result
	If result=0 Then Exit Function
	mailsubject=rs("MailSubject")
	mailcontent=rs("MailContent")
	arr_fieldname=Split(rs("FieldList"),",")
	CloseRs rs
	CloseConn conn
	' 应用变量值到邮件标题及内容中
	ReDim Preserve arr_fieldval(UBound(arr_fieldname))
	For i=0 To UBound(arr_fieldname)
		If Not IsNull(arr_fieldval(i)) Then
			mailsubject=Replace(mailsubject,arr_fieldname(i),arr_fieldval(i))
			mailcontent=Replace(mailcontent,arr_fieldname(i),arr_fieldval(i))
		End If
	Next
	mailsubject="【工管院务系统】"&mailsubject
	
	If Right(mailrcpt,12)<>"@scut.edu.cn" Then
		' 转换教育网域名为公网域名
		mailcontent=Replace(mailcontent,"edu.cnsba.com","pub.cnsba.com")
	End If
	If Len(mailrcpt) Then
		' 落款
		mailbody=mailcontent&"<p align=""right"">华南理工大学工商管理学院电子院务平台<br/>"&Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日<br/>(此邮件为系统自动发出，请勿回复)</p>"
		' Message组件发送邮件更快，但是无法清楚是否发送成功
'		Set jmail = Server.CreateObject("jmail.Message")
'		jmail.Charset = "utf-8" ' 邮件字符集，默认为"US-ASCII"
'		' jmail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
'		' 发送者信息（可用变量方式赋值）
'		jmail.From = "tutorsystem@cnsba.com" ' 发送者地址
'		jmail.FromName = "华南理工大学工商管理学院电子院务平台" ' 发送者姓名
'		jmail.Subject = mailsubject ' 邮件主题
'		' 身份验证
'		jmail.MailServerUserName = "" ' 身份验证的用户名
'		jmail.MailServerPassword = "" ' 身份验证的密码
'		' 设置优先级，范围从1到5，越大的优先级越高，3为普通
'		jmail.Priority = 1
'		jmail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
'		' 加入一个收件人【变量email：收件人地址】可以同一语句重复加入多个
'		jmail.AddRecipient(mailrcpt)
'		' 加入附件【变量filename：附件文件的绝对地址，确保用户IUSR_机器名有访问的权限】
'		' 【参数设置是(True)否(False)为Inline方式】
'		'contentId = jmail.AddAttachment (Server.MapPath("jmail.asp"), True)
'		' 邮件主体（HTML(注意信件内链接附件的方式)）
'		'jmail.AppendBodyFromFile(Server.MapPath("/Love/Inc/Mailend.txt"))
'		jmail.HTMLBody = mailbody
'		' 邮件主体（文本部分）
'		jmail.Body = mailsubject
'		' 发送【调用格式：objjmail.Send([username:password@]SMTPServerAddress[:Port])】
'		jmail.Send("202.38.194.226")
'		' 关闭并清除对象
'		jmail.Close()
'		Set jmail = Nothing
		' SMTPMail组件发送邮件显得更慢，可通过报错提示发送失败
		Set jmail=Server.CreateObject("JMail.SMTPMail")
		Dim attempt_num:attempt_num=0
		Dim max_attempt_num:max_attempt_num=1
		Do
			attempt_num=attempt_num+1
			jmail.Silent=True
			jmail.Logging=True
			jmail.Charset="utf-8"
			jmail.ContentType="text/html"
			jmail.AddRecipient mailrcpt
			jmail.Sender="tutorsystem@cnsba.com"
			jmail.Priority=1
			jmail.Subject=mailsubject
			jmail.Body=mailbody
			bSuccess=jmail.Execute()
		Loop While Not bSuccess And attempt_num<max_attempt_num
		jmail.Close()
		Set jmail=Nothing
	Else
		bSuccess=False
	End If
	sendAnnouncementEmail=bSuccess
End Function
Function sendSMS(mail_id,rcpt,arr_fieldval)
	If IsNull(mail_id) Then Exit Function
	Dim sql,rs,conn,result,i
	Dim arr_fieldname
	Dim sms,mailcontent,mailbody
	Dim ret
	' 查询邮件模板内容、变量等信息
	sql="SELECT * FROM EMAIL_TEMPLATE WHERE ID="&mail_id
	Connect conn
	GetRecordSetNoLock conn,rs,sql,result
	If result=0 Then Exit Function
	mailcontent=rs("MailContent")
	arr_fieldname=Split(rs("FieldList"),",")
	CloseRs rs
	CloseConn conn
	' 应用变量值到短信内容中
	ReDim Preserve arr_fieldval(UBound(arr_fieldname))
	For i=0 To UBound(arr_fieldname)
		If Not IsNull(arr_fieldval(i)) Then
			mailcontent=Replace(mailcontent,arr_fieldname(i),arr_fieldval(i))
		End If
	Next
	If Len(rcpt) Then
		mailbody=toPlainText(mailcontent&"(本短信为系统自动发出，请勿回复)")
		'On Error Resume Next
		' 发送请求到WebService
		Set sms=Server.CreateObject("SmsSender.SmsService")
		ret=sms.sendSMS(rcpt,mailbody)
		Set sms=Nothing
	Else
		ret=1
	End If
	sendSMS=ret
End Function

Function sendCustomEmail(mailrcpt,subject,content)
	Dim jmail,mailbody,bSuccess
	subject="【工管院务系统】"&subject
	If Right(mailrcpt,12)<>"@scut.edu.cn" Then
		' 转换教育网域名为公网域名
		content=Replace(content,"edu.cnsba.com","pub.cnsba.com")
	End If
	If Len(mailrcpt) Then
		' 落款
		mailbody=content&"<p align=""right"">华南理工大学工商管理学院电子院务平台<br/>"&Year(Now)&"年"&Month(Now)&"月"&Day(Now)&"日<br/>(此邮件为系统自动发出，请勿回复)</p>"
		Set jmail=Server.CreateObject("JMail.SMTPMail")
		jmail.silent=False
		jmail.logging=True
		jmail.Charset="utf-8"
		jmail.ContentType="text/html"
		jmail.AddRecipient mailrcpt
		jmail.Sender="tutorsystem@cnsba.com"
		jmail.Priority=1
		jmail.Subject=subject
		jmail.Body=mailbody
		bSuccess=jmail.Execute()
		jmail.Close()
		Set jmail=Nothing
	Else
		bSuccess=False
	End If
	sendCustomEmail=bSuccess
End Function
Function sendCustomSMS(rcpt,content)
	Dim mailbody,sms,ret
	If Len(rcpt) Then
		mailbody=toPlainText(content&"(本短信为系统自动发出，请勿回复)")
		'On Error Resume Next
		' 发送请求到WebService
		Set sms=Server.CreateObject("SmsSender.SmsService")
		ret=sms.sendSMS(rcpt,mailbody)
		Set sms=Nothing
	Else
		ret=-1
	End If
	sendCustomSMS=ret
End Function

Function getTutorSystemMailIdByType(d)
	Dim sem_info,curyear,cur_semester,mail_id(4),i
	Dim conn,rs,sql
	sem_info=getCurrentSemester()
	curyear=sem_info(0)
	cur_semester=sem_info(1)
	Connect conn
	sql="SELECT * FROM TUTOR_SYSTEM_SETTINGS WHERE USE_YEAR="&curyear&" AND USE_SEMESTER="&cur_semester
	Set rs=conn.Execute(sql)
	If Not rs.EOF Then
		For i=1 To UBound(mail_id)
			mail_id(i)=rs("MAIL_"&i)
		Next
	Else
		For i=1 To UBound(mail_id)
			mail_id(i)=Null
		Next
	End If
	CloseRs rs
	CloseConn conn
	getTutorSystemMailIdByType=mail_id
End Function
Function getThesisReviewSystemMailIdByType(d)
	Dim sem_info,curyear,cur_semester,mail_id(11),i
	Dim conn,rs,sql
	sem_info=getCurrentSemester()
	curyear=sem_info(0)
	cur_semester=sem_info(1)
	Connect conn
	sql="SELECT * FROM TEST_THESIS_REVIEW_SYSTEM WHERE USE_YEAR="&curyear&" AND USE_SEMESTER="&cur_semester
	Set rs=conn.Execute(sql)
	If Not rs.EOF Then
		For i=1 To UBound(mail_id)
			mail_id(i)=rs("MAIL_"&i)
		Next
	Else
		For i=1 To UBound(mail_id)
			mail_id(i)=Null
		Next
	End If
	CloseRs rs
	CloseConn conn
	getThesisReviewSystemMailIdByType=mail_id
End Function
%>