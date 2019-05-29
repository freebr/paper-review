<%
'==========================================
' API 名称：    remove-activity
' API 功能：    删除评阅活动
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-03
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main(args)
    Dim data: Set data=CreateDictionary()
    Dim arg: Set arg=CreateDictionary()
    ensureArgument args, arg, data
    Dim conn,sql,count
    Connect conn
    sql="DECLARE @Id int=?;UPDATE Activities SET Valid=0 WHERE Id=@Id;"&_
    "DELETE FROM ActivityPeriods WHERE ActivityId=@Id;"&_
    "DELETE FROM ActivityMailTemplates WHERE ActivityId=@Id"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,CmdParam("Id",adInteger,4,arg("id")))
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

Call main("id")
%>