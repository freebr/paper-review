<%
'==========================================
' API 名称：    set-mail-template
' API 功能：    设置指定的通知邮件（短信）模板
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-05
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main(args)
    Dim data: Set data=CreateDictionary()
    Dim arg: Set arg=CreateDictionary()
    ensureArgument args, arg, data
    Dim conn,sql,count
    ConnectDb conn
    sql="UPDATE ActivityMailTemplates SET MailContent=? WHERE Id=?"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,_
        CmdParam("MailContent",adVarWChar,3000,arg("mail_content")),_
        CmdParam("Id",adInteger,4,arg("id")))
    If Err.Number Then
        data.Add "status", "error"
        data.Add "msg", Err.Description
    Else
        data.Add "status", "ok"
    End If
    On Error GoTo 0
    CloseConn conn

    Call (new JSONWriter)(Response, data)
End Function

Call main(Array("id", "mail_content"))
%>