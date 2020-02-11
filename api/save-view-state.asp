<%
'==========================================
' API 名称：    save-view-state
' API 功能：    保存视图状态
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
    On Error Resume Next
	setViewState arg("user_id"), arg("user_type"), arg("view_name"), arg("view_state")    
    If Err.Number Then
        data.Add "status", "error"
        data.Add "msg", Err.Description
    Else
        data.Add "status", "ok"
    End If
    On Error GoTo 0
    
    Call (new JSONWriter)(Response, data)
End Function

Call main(Array("user_id", "user_type", "view_name", "view_state"))
%>