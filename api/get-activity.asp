<%
'==========================================
' API 名称：    get-activity
' API 功能：    提供已录入的单个评阅活动信息
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
    Dim conn,rs,sql,count
    Connect conn
    sql="SELECT Id,Name,SemesterId,SemesterName,StuTypeId,StuTypeName,IsOpen,CreatedAt,Creator,CreatorName FROM ViewActivities WHERE Id=?"
    Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("ActivityId",adInteger,4,arg("id")))
    Set rs=ret("rs")
    count=ret("count")

    Dim dictItem: Set dictItem=CreateDictionary()
    Dim i
    For i=0 To rs.Fields.Count-1
        dictItem.Add rs.Fields(i).Name, rs(i).Value
    Next
    data.Add "status", "ok"
    data.Add "data", dictItem

    CloseRs rs
    CloseConn conn
    
    Call (new JSONWriter)(Response, data)
End Function

Call main("id")
%>