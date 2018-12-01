﻿<%Response.Charset="utf-8"%>
<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
stu_type=Session("StuType")

Dim arrTemplateNames:arrTemplateNames=Array("开题报告表","中期检查表","预答辩意见书","审批材料表","硕士学位论文送审申请表","硕士学位论文文字复制比情况说明表","硕士学位论文分会复审意见表")
Dim arrTemplateFiles:arrTemplateFiles=Array("ktbg.doc","zqjcb.doc","ydbyjs.doc","spcl.doc","sssqb.doc","fzbsmb.doc","fsyjb.doc")
prefix0="new/"
Select Case stu_type
Case 5:prefix=prefix0&"me_"
Case 6:prefix=prefix0&"mba_"
Case 7:prefix=prefix0&"emba_"
Case 9:prefix=prefix0&"mpacc_"
End Select
Dim arrFileIndexToAddPrefix:arrFileIndexToAddPrefix=Array(0,1,3)
For i=0 To UBound(arrFileIndexToAddPrefix)
	arrTemplateFiles(arrFileIndexToAddPrefix(i))=prefix&arrTemplateFiles(arrFileIndexToAddPrefix(i))
Next

Dim arrSpecMatNames:arrSpecMatNames=Array("研究生学位论文撰写规范","MBA论文撰写手册","MPAcc论文撰写手册") ',"2017年MBA导师分组表及联系方式")
Dim arrSpecMatUsers:arrSpecMatUsers=Array("*","6","9")
Dim arrSpecMatFiles:arrSpecMatFiles=Array("lwzxgf.doc",prefix0&"mba_lwzxsc20170714.pdf",prefix0&"mpacc_lwzxsc20170713.pdf",prefix0&"mba_dsfzb.pdf")

%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../css/student.css" rel="stylesheet" type="text/css" />
<script src="../scripts/jquery-1.11.3.min.js" type="text/javascript"></script>
<script src="../scripts/utils.js" type="text/javascript"></script>
<meta name="theme-color" content="#2D79B2" />
<title>查看论文阶段相关资料</title>
</head>
<body bgcolor="ghostwhite">
<center><font size=4><b>查看论文阶段相关资料</b></font>
<table class="tblform" width="800"><tr><td align="left"><%
For i=0 To UBound(arrTemplateNames)
	link="template/doc/"&arrTemplateFiles(i)
	ext=UCase(Mid(arrTemplateFiles(i),InStrRev(arrTemplateFiles(i),".")+1))
%><p><a href="<%=link%>" target="_blank" title="<%=ext%>格式"><img src="../images/student/<%=ext%>.png" width="16" height="16">下载<%=arrTemplateNames(i)%></a></p><%
Next

For i=0 To UBound(arrSpecMatNames)
	link="template/doc/"&arrSpecMatFiles(i)
	If arrSpecMatUsers(i)="*" Or InStr(arrSpecMatUsers(i),stu_type) Then
		ext=UCase(Mid(arrSpecMatFiles(i),InStrRev(arrSpecMatFiles(i),".")+1))
%><p><a href="<%=link%>" target="_blank" title="<%=ext%>格式"><img src="../images/student/<%=ext%>.png" width="16" height="16">下载<%=arrSpecMatNames(i)%></a></p><%
	End If
Next
%>
</td></tr></table></center></body></html>