<%
'==========================================
' API 名称：    get-audit-records
' API 功能：    提供指定论文的审核记录
' API 输出类型：JSON
' 修订日期：    2020-04-07
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main(args)
    Dim data: Set data=CreateDictionary()
    Dim arg: Set arg=CreateDictionary()
    ensureArgument args, arg, data
    Dim arr:arr=getAuditInfo(arg("id"), Null, Null)
    data.Add "status", "ok"
    data.Add "data", arr

    CloseRs rs
    CloseConn conn
    
    Call (new JSONWriter)(Response, data)
End Function

Call main("id")
%>