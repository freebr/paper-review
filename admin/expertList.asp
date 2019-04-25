<!--#include file="../inc/db.asp"-->
<!--#include file="common.asp"-->
<%If IsEmpty(Session("Id")) Then Response.Redirect("../error.asp?timeout")
Dim finalFilter,pageNo,pageSize
pageNo=""
pageSize=""
If Request.Form("finalFilter2").Count=0 Then
	finalFilter=Request.Form("finalFilter")
	pageSize=Request.Form("pageSize")
	pageNo=Request.Form("pageNo")
Else
	finalFilter=Request.Form("finalFilter2")
	pageSize=Request.Form("pageSize2")
	pageNo=Request.Form("pageNo2")
End If
If Len(finalFilter) Then PubTerm="AND ("&finalFilter&")"
Connect conn
sql="SELECT * FROM ViewExpertInfo WHERE Valid=1 "&PubTerm&" ORDER BY EXPERT_NAME"
GetRecordSetNoLock conn,rs,sql,result
If IsEmpty(pageSize) Or Not IsNumeric(pageSize) Then
  pageSize=-1
Else
	pageSize=CInt(pageSize)
End If
If pageSize=-1 Then
	If rs.RecordCount>0 Then rs.PageSize=rs.RecordCount
Else
  rs.PageSize=pageSize
End If
If IsEmpty(pageNo) Or Not IsNumeric(pageNo) Then
	If rs.PageCount<>0 Then pageNo=1
Else
	pageNo=CInt(pageNo)
  If pageNo>rs.PageCount Then
	  If rs.PageCount<>0 Then pageNo=1
	End If
End If
If rs.RecordCount>0 Then rs.AbsolutePage=pageNo
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<% useStylesheet("admin") %>
<% useScript(Array("jquery", "common", "upload", "expertList")) %>
</head>
<body bgcolor="ghostwhite">
<center>
<font size=4><b>专业硕士论文评阅专家名单</b></font>
<form id="fmUpload" action="importExpertInfo.asp?step=2" method="POST" enctype="multipart/form-data">
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
<tr bgcolor="ghostwhite"><td style="font-weight:bold">从Excel文件导入评阅专家信息&emsp;<a href="upload/exp_template.xlsx" target="_blank">点击下载专家信息表格模板</a></td></tr>
<tr bgcolor="ghostwhite"><td>
<p>请选择要导入的 Excel 文件：<br />文件名：<input type="file" name="excelFile" size="100" />
<input type="submit" name="btnsubmit" value="提 交" />&nbsp;</p>
</td></tr></table>
</form>
<table cellspacing=4 cellpadding=0>
<form id="query" method="post" onsubmit="return chkField()">
<tr><td>
<!--查找-->
<select name="field" onchange="ReloadOperator()">
<option value="s_EXPERT_NAME">姓名</option>
<option value="s_EXPERTISE">学科专长</option>
<option value="s_WORKPLACE">单位（住址）</option>
</select>
<select name="operator">
<script>ReloadOperator()</script>
</select>
<input type="text" name="filter" size="10" onkeypress="checkKey()">
<input type="hidden" name="finalFilter" value="<%=finalFilter%>">
<input type="submit" value="查找" onclick="genFilter()">
<input type="submit" value="在结果中查找" onclick="genFinalFilter()">
&nbsp;
每页
<select name="pageSize" onchange="this.form.submit()">
<option value="-1" <%If pageSize=-1 Then%>selected<%End If%>>全部</option>
<option value="20" <%If rs.PageSize=20 Then%>selected<%End If%>>20</option>
<option value="40" <%If rs.PageSize=40 Then%>selected<%End If%>>40</option>
<option value="60" <%If rs.PageSize=60 Then%>selected<%End If%>>60</option>
</select>
条
&nbsp;
转到
<select name="pageNo" onchange="this.form.submit()">
<%
For i=1 to rs.PageCount
    Response.write "<option value="&i
    If rs.AbsolutePage=i Then Response.write " selected"
    Response.write ">"&i&"</option>"
Next
%>
</select>
页
&nbsp;
共<%=rs.RecordCount%>条
</td><td>全选<input type="checkbox" onclick="checkAll()" id="chk" />&nbsp;<input type="button" value="删 除" onclick="if(confirm('是否删除这'+countClk()+'条记录？'))$('#fmExpList').submit()" /></td></tr>
<tr><td colspan="2" align="center">&nbsp;<input type="button" id="btnresetpwd" value="重置账号密码" />
&nbsp;<input type="button" value="批量发送短信" onclick="batchSendNotice($('#fmExpList'),1)" />
&nbsp;<input type="button" value="批量发送邮件" onclick="batchSendNotice($('#fmExpList'),0)" />
&nbsp;<input type="button" id="btnexport" value="导出到Excel文件" /></td></tr></form></table>
<form id="fmExpList" method="post" action="delExpert.asp">
<input type="hidden" name="finalFilter2" value="<%=finalFilter%>" />
<input type="hidden" name="pageSize2" value=<%=pageSize%> />
<input type="hidden" name="pageNo2" value=<%=pageNo%> />
<input type="hidden" name="batch" value="1" />
<table width="1000" cellpadding="2" cellspacing="1" bgcolor="dimgray">
  <tr bgcolor="gainsboro" align="center" height="25">
    <td width="160">姓名/登录名</td>
    <td width="80">职称</td>
    <td width="80">最高学历</td>
    <td width="100">学科专长</td>
		<td>单位（住址）</td>
		<td width="100">联系电话</td>
		<td width="120">邮箱</td>
		<td width="120">操作</td>
    <td width="30">选择</td>
  </tr>
  <%
  Dim teacherno,last_diploma,bSelectable
  For i=1 to rs.PageSize
      If rs.EOF Then Exit For
      teacherno=rs("TEACHERNO").Value
      last_diploma=rs("LAST_DIPLOMA").Value
      If IsNull(last_diploma) Then last_diploma=0
      bSelectable=teacherno<>"zhuanjia1" And teacherno<>"zhuanjia2"
  %>
  <tr bgcolor="ghostwhite">
    <td align=center><a href="expertProfile.asp?id=<%=rs("TEACHER_ID").Value%>"><%=HtmlEncode(rs("EXPERT_NAME").Value)%>&nbsp;/&nbsp;<%=HtmlEncode(teacherno)%></a></td>
    <td align=center><%=HtmlEncode(rs("PRO_DUTY_NAME").Value)%></td>
    <td align=center><%=arrDiplomaName(last_diploma)%></td>
    <td align=center><%=HtmlEncode(rs("EXPERTISE").Value)%></td>
    <td align=center><%=HtmlEncode(rs("ADDRESS").Value)%></td>
    <td align=center><%=HtmlEncode(rs("MOBILE").Value)%></td>
    <td align=center><%=HtmlEncode(rs("EMAIL").Value)%></td>
    <td align=center><a id="pwd<%=i%>" href="#" onclick="return showPassword(this,'<%=rs("PASSWORD").Value%>')">显示密码</a>
    <a href="expertProfile.asp?id=<%=rs("TEACHER_ID").Value%>">查看资料</a><br/><a href="#" onclick="window.open('sendmsg.asp?type=1&tid=<%=rs("TEACHER_ID").Value%>','','width=1010,height=420,status=no');return false">发送短信</a>&nbsp;<a href="#" onclick="window.open('sendmsg.asp?type=2&tid=<%=rs("TEACHER_ID").Value%>','','width=1010,height=420,status=no');return false">发送邮件</a></td>
    <td align=center><%
    	If bSelectable Then
    %><input type="checkbox" name="sel" value="<%=rs("TEACHER_ID").Value%>"><input type="hidden" name="isinschool<%=rs("TEACHER_ID").Value%>" value="<%=Abs(rs("INSCHOOL").Value)%>"><%
  		End If %>
	</td></tr>
  <%
  	rs.MoveNext()
  Next
%></table></center></body>
<script type="text/javascript">
	$(document).ready(function(){
    $('#fmUpload :submit').click(function() {
      var valid=checkIfExcel(this.excelFile);
      if(valid) {
        $(':submit').val("正在提交，请稍候...").attr('disabled',true);
      }
      return valid;
    }).attr('disabled',false);
    $('#btnresetpwd').click(function() {
      $(this).val('正在处理，请稍候……').attr('disabled',true);
      resetPassword($('#fmExpList'));
    }).attr('disabled',false);
    $('#btnexport').click(function() {
      $(this).val('正在导出，请稍候……').attr('disabled',true);
      exportInfo($('#fmExpList'));
    }).attr('disabled',false);
	});
</script></html><%
  CloseRs rs
  CloseConn conn
%>