<%
'==========================================
' API 名称：    set-system-status
' API 功能：    设置系统开放状态
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
	sql="UPDATE Configs SET Status=?, StatusUpdateTime=GETDATE()"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,CmdParam("Status",adVarWChar,50,arg("status")))
    If Err.Number Then
        data.Add "status", "error"
        data.Add "msg", Err.Description
    Else
        data.Add "status", "ok"

        Select Case arg("status")
        Case "closed"
            writeLog "行政人员["&Session("name")&"]设置系统为关闭状态。"
        Case "open"
            writeLog "行政人员["&Session("name")&"]设置系统为开放状态。"
        Case "debug"
            writeLog "行政人员["&Session("name")&"]设置系统为调试状态。"
        End Select
    End If
    On Error GoTo 0
    CloseConn conn

    Call (new JSONWriter)(Response, data)
End Function

Call main("status")
%>