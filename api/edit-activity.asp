<%
'==========================================
' API 名称：    edit-activity
' API 功能：    编辑评阅活动
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-09
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
    sql="UPDATE Activities SET Name=?,SemesterId=?,IsOpen=? WHERE Id=?"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,_
        CmdParam("Name",adVarWChar,50,arg("name")),_
        CmdParam("SemesterId",adInteger,4,arg("semester")),_
        CmdParam("IsOpen",adBoolean,1,arg("is_open")),_
        CmdParam("Id",adInteger,4,arg("id")))
    If Err.Number Then
        data.Add "status", "error"
        data.Add "msg", Err.Description
    Else
        data.Add "status", "ok"
    End If
    On Error GoTo 0
    CloseRs rs
    CloseConn conn

    Call (new JSONWriter)(Response, data)
End Function

Call main(Array("id", "name", "semester", "is_open"))
%>