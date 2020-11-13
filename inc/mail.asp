<!--#include file="sms.inc"--><%

Function sendNotifyMail(activity_id,stu_type_id,template_name,recipient,field_values)
	Dim conn:ConnectDb conn
	' 获取邮件模板信息
	Dim sql:sql="SELECT * FROM ActivityMailTemplates WHERE ActivityId=? AND StuType=? AND Name=?"
	Dim ret:Set ret=ExecQuery(conn,sql,_
		CmdParam("ActivityId",adInteger,4,activity_id),_
		CmdParam("StuType",adInteger,4,stu_type_id),_
		CmdParam("Name",adVarWChar,50,template_name))
	If ret("count")=0 Then Exit Function

	Dim rs:Set rs=ret("rs")
	Dim mail_subject:mail_subject=rs("MailSubject")
	Dim mail_content:mail_content=rs("MailContent")
	CloseRs rs
	CloseConn conn

	' 用实际值替换邮件标题及内容中的符号
	For Each name In field_values
		Dim value:value=field_values(name)
		If IsNull(value) Then value=""
		mail_subject=Replace(mail_subject,"$"&name,value)
		mail_content=Replace(mail_content,"$"&name,value)
	Next
	mail_subject="【工管院务系统】"&mail_subject

	Dim is_sent
	If Len(recipient) Then
		' 增加落款
		Dim mail_body:mail_body=mail_content&_
			Format("<p align=""right"">华南理工大学工商管理学院电子院务平台<br/>{0}<br/>(此邮件为系统自动发出，请勿回复)</p>",_
				FormatDateTime(Now(),1))
		' SMTPMail组件发送邮件显得更慢，可通过报错提示发送失败
		Dim jmail:Set jmail=Server.CreateObject("JMail.SMTPMail")
		jmail.Silent=True
		jmail.Logging=True
		jmail.Charset="utf-8"
		jmail.ContentType="text/html"
		jmail.AddRecipient recipient
		jmail.Sender="lunwenxitong@cnsba.com"
		jmail.Priority=1
		jmail.Subject=mail_subject
		jmail.Body=mail_body

		Dim attempt_num:attempt_num=0
		Dim max_attempt_num:max_attempt_num=1
		Do
			attempt_num=attempt_num+1
			is_sent=jmail.Execute()
		Loop While Not is_sent And attempt_num<max_attempt_num
		jmail.Close()
		Set jmail=Nothing
	Else
		is_sent=False
	End If
	sendNotifyMail=is_sent
End Function

Function sendNotifySms(activity_id,stu_type_id,template_name,recipient,field_values)
	Dim conn:ConnectDb conn
	' 获取短信模板信息
	Dim sql:sql="SELECT * FROM ActivityMailTemplates WHERE ActivityId=? AND StuType=? AND Name=?"
	Dim ret:Set ret=ExecQuery(conn,sql,_
		CmdParam("ActivityId",adInteger,4,activity_id),_
		CmdParam("StuType",adInteger,4,stu_type_id),_
		CmdParam("Name",adVarWChar,50,template_name))
	If ret("count")=0 Then Exit Function

	Dim rs:Set rs=ret("rs")
	Dim sms_content:sms_content=rs("MailContent")
	CloseRs rs
	CloseConn conn

	' 用实际值替换短信内容中的符号
	For Each name In field_values
		If Not IsNull(field_values(name)) Then
			sms_content=Replace(sms_content,"$"&name,field_values(name))
		End If
	Next

	If Len(recipient) Then
		Dim sms_body:sms_body=toPlainText(sms_content&"(此短信为系统自动发出，请勿回复)")
		Dim msg:Set msg=New Messenger
		msg.MsgType=0
		ret=msg.sendMessage(recipient, sms_body)
		Set msg=Nothing
	Else
		ret=1
	End If
	sendNotifySms=ret=0
End Function

Function sendCustomEmail(recipient,ByVal subject,ByVal content)
	Dim is_sent
	subject="【工管院务系统】"&subject
	If Len(recipient) Then
		' 落款
		Dim mail_body:mail_body=content&_
			Format("<p align=""right"">华南理工大学工商管理学院电子院务平台<br/>{0}<br/>(此邮件为系统自动发出，请勿回复)</p>",_
				FormatDateTime(Now(),1))
		Dim jmail:Set jmail=Server.CreateObject("JMail.SMTPMail")
		jmail.silent=False
		jmail.logging=True
		jmail.Charset="utf-8"
		jmail.ContentType="text/html"
		jmail.AddRecipient recipient
		jmail.Sender="lunwenxitong@cnsba.com"
		jmail.Priority=1
		jmail.Subject=subject
		jmail.Body=mail_body
		is_sent=jmail.Execute()
		jmail.Close()
		Set jmail=Nothing
	Else
		is_sent=False
	End If
	sendCustomEmail=is_sent
End Function

Function sendCustomSms(recipient,content)
	Dim ret
	If Len(recipient) Then
		Dim sms_body:sms_body=toPlainText(content&"(本短信为系统自动发出，请勿回复)")
		Dim msg:Set msg=New Messenger
		msg.MsgType=0
		ret=msg.sendMessage(recipient, sms_body)
		Set msg=Nothing
	Else
		ret=-1
	End If
	sendCustomSms=ret=0
End Function

'		' Message组件发送邮件更快，但是不能确定发送状态
'		Set jmail = Server.CreateObject("jmail.Message")
'		jmail.Charset = "utf-8" ' 邮件字符集，默认为"US-ASCII"
'		' jmail.ISOEncodeHeaders = False ' 是否进行ISO编码，默认为True
'		' 发送者信息（可用变量方式赋值）
'		jmail.From = "lunwenxitong@cnsba.com" ' 发送者地址
'		jmail.FromName = "华南理工大学工商管理学院电子院务平台" ' 发送者姓名
'		jmail.Subject = mail_subject ' 邮件主题
'		' 身份验证
'		jmail.MailServerUserName = "" ' 身份验证的用户名
'		jmail.MailServerPassword = "" ' 身份验证的密码
'		' 设置优先级，范围从1到5，越大的优先级越高，3为普通
'		jmail.Priority = 1
'		jmail.AddHeader "Originating-IP", Request.ServerVariables("REMOTE_ADDR")
'		' 加入一个收件人【变量email：收件人地址】可以同一语句重复加入多个
'		jmail.AddRecipient(recipient)
'		' 加入附件【变量filename：附件文件的绝对地址，确保用户IUSR_机器名有访问的权限】
'		' 【参数设置是(True)否(False)为Inline方式】
'		'contentId = jmail.AddAttachment (Server.MapPath("jmail.asp"), True)
'		' 邮件主体（HTML(注意信件内链接附件的方式)）
'		'jmail.AppendBodyFromFile(Server.MapPath("/Love/Inc/Mailend.txt"))
'		jmail.HTMLBody = mail_body
'		' 邮件主体（文本部分）
'		jmail.Body = mail_subject
'		' 发送【调用格式：objjmail.Send([username:password@]SMTPServerAddress[:Port])】
'		jmail.Send("202.38.194.226")
'		' 关闭并清除对象
'		jmail.Close()
'		Set jmail = Nothing
%>