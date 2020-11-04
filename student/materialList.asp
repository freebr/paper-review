<!--#include file="../inc/global.inc"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("StuId")) Then Response.Redirect("../error.asp?timeout")
stu_type=Session("StuType")
version="20200814"

Dim dictCommonMat:Set dictCommonMat=CreateDictionary()
dictCommonMat.Add "开题报告表","ktbgb.doc"
dictCommonMat.Add "中期考核表","zqkhb.doc"
dictCommonMat.Add "预答辩意见书","ydbyjs.doc"
dictCommonMat.Add "审批材料表（适用于仅申请学位者）","spclb_2.doc"
dictCommonMat.Add "审批材料表（适用于申请毕业及学位者）","spclb.doc"
dictCommonMat.Add "硕士学位论文送审申请表","sssqb.doc"
dictCommonMat.Add "硕士学位论文文字复制比情况说明表","fzbsmb.doc"
dictCommonMat.Add "硕士学位论文分会复审意见表","fsyjb.doc"

Dim dictSpecMat:Set dictSpecMat=CreateDictionary()
dictSpecMat.Add "研究生学位论文撰写规范",Array("*","lwzxgf.doc")
dictSpecMat.Add "Materials for Verification and Approval of Master’s Degree Dissertation (MBA) Defence and Degree Conferral",Array("6","mba_spclb_en.doc")
dictSpecMat.Add "MBA论文撰写手册",Array("6","mba_lwzxsc20170714.pdf")
dictSpecMat.Add "MPAcc论文撰写手册",Array("9","mpacc_lwzxsc20170713.pdf")
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="theme-color" content="#2D79B2" />
<title>查看论文阶段相关资料</title>
<% useStylesheet "student" %>
<% useScript "jquery", "common" %>
</head>
<body>
<center><font size=4><b>查看论文阶段相关资料</b></font>
<table class="form" width="800"><tr><td align="left"><%
For Each name In dictCommonMat
	link=resolvePath("template/doc/"&version,dictCommonMat(name))
	ext=UCase(Mid(dictCommonMat(name),InStrRev(dictCommonMat(name),".")+1))
%><p><a href="<%=link%>" target="_blank" title="<%=ext%>格式"><img src="../images/student/<%=ext%>.png" width="16" height="16">下载<%=name%></a></p><%
Next

For Each name In dictSpecMat
	link=resolvePath("template/doc/"&version,dictSpecMat(name)(1))
	If dictSpecMat(name)(0)="*" Or InStr(dictSpecMat(name)(0),stu_type) Then
		ext=UCase(Mid(dictSpecMat(name)(1),InStrRev(dictSpecMat(name)(1),".")+1))
%><p><a href="<%=link%>" target="_blank" title="<%=ext%>格式"><img src="../images/student/<%=ext%>.png" width="16" height="16">下载<%=name%></a></p><%
	End If
Next
%>
</td></tr></table></center></body></html>