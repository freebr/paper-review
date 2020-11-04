<%
'==========================================
' API 名称：    delete-review-record
' API 功能：    删除评阅记录
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2020-5-11
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
    sql="DELETE FROM ReviewRecords WHERE Id=?"
    'On Error Resume Next
    count=ExecNonQuery(conn,sql,CmdParam("Id",adGUID,16,arg("id")))
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

Call main(Array("id"))
%>