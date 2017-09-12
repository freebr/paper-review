<%Response.Charset="utf-8"
Response.Expires=-1%>
<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/upload_5xsoft.inc"-->
<%If IsEmpty(Session("user")) Then Response.Redirect("../error.asp?timeout")
Set upload=New upload_5xsoft
TeacherId=upload.Form("teacherid")
FormGetToSafeRequest(TeacherId)
If Len(TeacherId)=0 Or Not IsNumeric(TeacherId) Then
%><body bgcolor="ghostwhite"><center><font color=red size="4">参数无效。</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	Response.End
End If

teachername=upload.Form("teachername")
teacherno=upload.Form("teacherno")
sex=upload.Form("sex")
pro_duty_name=upload.Form("pro_duty_name")
expertise=upload.Form("expertise")
email=upload.Form("email")
workplace=upload.Form("workplace")
address=upload.Form("address")
mailcode=upload.Form("mailcode")
telephone=upload.Form("telephone")
mobile=upload.Form("mobile")
bankaccount=upload.Form("bankaccount")
bankname=upload.Form("bankname")
idcard_no=upload.Form("idcard_no")

Connect conn
sql="SELECT * FROM SCUT_MD.dbo.TEACHER_INFO WHERE TEACHERID="&Teacherid
GetRecordSet conn,rs,sql,result
If sex<>"男" And sex<>"女" Then
	bError=True
	errdesc="请选择性别！"
ElseIf Len(teachername)=0 Then
	bError=True
	errdesc="请填写姓名！"
ElseIf Len(teacherno)=0 Then
	bError=True
	errdesc="请填写登录名！"
ElseIf Len(pro_duty_name)=0 Then
	bError=True
	errdesc="请填写专业技术职务（职称）！"
ElseIf Len(expertise)=0 Then
	bError=True
	errdesc="请填写学科专长！"
ElseIf Len(email)=0 Then
	bError=True
	errdesc="请填写电子邮箱！"
ElseIf Len(workplace)=0 Then
	bError=True
	errdesc="请填写单位名称！"
ElseIf Len(address)=0 Then
	bError=True
	errdesc="请填写通信地址！"
ElseIf Len(address)>25 Then
	bError=True
	errdesc="通信地址最多只能填25字！"
ElseIf Len(mailcode)=0 Then
	bError=True
	errdesc="请填写邮编！"
ElseIf Len(telephone)=0 Then
	bError=True
	errdesc="请填写联系电话（办公室）！"
ElseIf Len(mobile)=0 Then
	bError=True
	errdesc="请填写联系电话（移动）！"
ElseIf Len(bankaccount)=0 Then
	bError=True
	errdesc="请填写银行账户号！"
ElseIf Len(bankname)=0 Then
	bError=True
	errdesc="请填写开户行名称！"
ElseIf Len(idcard_no)=0 Then
	bError=True
	errdesc="请填写身份证号码！"
ElseIf rs.EOF Then
	bError=True
	errdesc="数据库没有记录！"
End If
If bError Then
%><body bgcolor="ghostwhite"><center><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></body><%
	CloseRs rs
  CloseConn conn
	Response.End
End If

If rs("IFTEACHER")=3 Then
	' 校外导师则更新教师信息表
	rs("TEACHERNO")=teacherno
	rs("TEACHERNAME")=teachername
	rs("SEX")=sex
	rs("OFFICE_ADDRESS")=workplace
	rs("TELPHONE")=telephone
	rs("MOBILE")=mobile
	rs("EMAIL")=email
	rs("IDCARD")=idcard_no
	rs.Update()
End If
CloseRs rs

' 更新专家库
sql="SELECT * FROM TEST_THESIS_REVIEW_EXPERT_INFO WHERE TEACHER_ID="&TeacherId
GetRecordSet conn,rs,sql,result
rs("EXPERT_NAME")=teachername
rs("PRO_DUTY_NAME")=pro_duty_name
rs("EXPERTISE")=expertise
rs("WORKPLACE")=workplace
rs("ADDRESS")=address
rs("MAILCODE")=mailcode
rs("TELEPHONE")=telephone
rs("MOBILE")=mobile
rs("EMAIL")=email
rs("BANK_ACCOUNT")=bankaccount
rs("BANK_NAME")=bankname
rs("IDCARD_NO")=idcard_no
rs.Update()
CloseRs rs
CloseConn conn
Set upload=Nothing
%><script type="text/javascript">
	alert("操作完成。");
	location.href="expertProfile.asp?id=<%=TeacherId%>";
</script>