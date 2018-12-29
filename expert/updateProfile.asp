<!--#include file="../inc/db.asp"-->
<!--#include file="../inc/ExtendedRequest.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("TId")) Then Response.Redirect("../error.asp?timeout")
TeacherId=Session("Tid")
If Len(TeacherId)=0 Or Not IsNumeric(TeacherId) Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4">参数错误。</font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></div></body></html><%
	Response.End()
End If

Set upload=New ExtendedRequest
teachername=upload.Form("teachername")
sex=upload.Form("sex")
pro_duty_name=upload.Form("pro_duty_name")
last_diploma=upload.Form("last_diploma")
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
newpwd=upload.Form("newpwd")
repeatpwd=upload.Form("repeatpwd")

Connect conn
ConnectOriginDb connOrigin
sql="SELECT * FROM TEACHER_INFO WHERE TEACHERID="&Teacherid
GetRecordSet connOrigin,rs,sql,result
If sex<>"男" And sex<>"女" Then
	bError=True
	errdesc="请选择性别！"
ElseIf Len(teachername)=0 Then
	bError=True
	errdesc="请填写姓名！"
ElseIf Len(pro_duty_name)=0 Then
	bError=True
	errdesc="请填写专业技术职务（职称）！"
ElseIf Len(last_diploma)=0 Or Not IsNumeric(last_diploma) Or last_diploma="0" Then
	bError=True
	errdesc="请选择最高学历！"
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
ElseIf newpwd<>repeatpwd Then
	bError=True
	errdesc="两次输入的密码不相同！"
ElseIf rs.EOF Then
	bError=True
	errdesc="数据库没有记录！"
End If
If bError Then
%><html><head><link href="../css/tutor.css" rel="stylesheet" type="text/css" /></head>
<body class="exp"><center><div class="content"><font color=red size="4"><%=errdesc%></font><br /><input type="button" value="返 回" onclick="history.go(-1)" /></center></div></body></html><%
	CloseRs rs
  CloseConn connOrigin
  CloseConn conn
	Response.End()
End If

If rs("IFTEACHER").Value=3 Then
	' 校外导师则更新教师信息表
	rs("TEACHERNAME").Value=teachername
	rs("SEX").Value=sex
	rs("Office_Address").Value=workplace
	rs("TELPHONE").Value=telephone
	rs("MOBILE").Value=mobile
	rs("EMAIL").Value=email
	rs("IDCARD").Value=idcard_no
End If
If Len(newpwd) Then
	rs("USER_PASSWORD").Value=newpwd
End If
rs.Update()
CloseRs rs
CloseConn connOrigin

' 更新专家库
sql="SELECT * FROM TEST_THESIS_REVIEW_EXPERT_INFO WHERE TEACHER_ID="&TeacherId
GetRecordSet conn,rs,sql,result
rs("EXPERT_NAME").Value=teachername
rs("PRO_DUTY_NAME").Value=pro_duty_name
rs("LAST_DIPLOMA").Value=last_diploma
rs("EXPERTISE").Value=expertise
rs("WORKPLACE").Value=workplace
rs("ADDRESS").Value=address
rs("MAILCODE").Value=mailcode
rs("TELEPHONE").Value=telephone
rs("MOBILE").Value=mobile
rs("EMAIL").Value=email
rs("BANK_ACCOUNT").Value=bankaccount
rs("BANK_NAME").Value=bankname
rs("IDCARD_NO").Value=idcard_no
rs.Update()
CloseRs rs
CloseConn conn
Set upload=Nothing
%><script type="text/javascript">
	alert("操作完成。");
	location.href="profile.asp";
</script>