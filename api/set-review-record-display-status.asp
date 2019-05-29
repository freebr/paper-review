<%
'==========================================
' API 名称：    set-review-record-display-status
' API 功能：    设置评阅记录显示状态
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-20
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
    sql="UPDATE ReviewRecords SET DisplayStatus=? WHERE Id=?"
    'On Error Resume Next
    count=ExecNonQuery(conn,sql,_
        CmdParam("DisplayStatus",adInteger,4,arg("display_status")),_
        CmdParam("Id",adGUID,16,arg("id")))
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

Call main(Array("id", "display_status"))
%>