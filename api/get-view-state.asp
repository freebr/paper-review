<%
'==========================================
' API 名称：    get-view-state
' API 功能：    获取视图状态
' API 输出类型：JSON
' 修订日期：    2019-12-22
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main(args)
    Dim data: Set data=CreateDictionary()
    Dim arg: Set arg=CreateDictionary()
    ensureArgument args, arg, data
    Dim state: state = getViewState(arg("user_id"), arg("user_type"), arg("view_name"))
    data.Add "status", "ok"
    data.Add "data", state

    CloseRs rs
    CloseConn conn
    
    Call (new JSONWriter)(Response, data)
End Function

Call main(Array("user_id", "user_type", "view_name"))
%>