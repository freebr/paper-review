<%
'==========================================
' API 名称：    set-activity-section-periods
' API 功能：    设置已录入的评阅活动某一环节的开放时间信息
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-01
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
    sql="UPDATE ActivityPeriods SET StartTime=?, EndTime=?, Enabled=? WHERE ActivityId=? AND StuType=? AND SectionId=?"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,_
        CmdParam("StartTime",adDate,4,arg("start_time")),_
        CmdParam("EndTime",adDate,4,arg("end_time")),_
        CmdParam("Enabled",adBoolean,1,arg("enabled")),_
        CmdParam("ActivityId",adInteger,4,arg("activity_id")),_
        CmdParam("StuType",adInteger,4,arg("stu_type")),_
        CmdParam("SectionId",adInteger,4,arg("section_id")))
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

Call main(Array("activity_id", "stu_type", "section_id", "start_time", "end_time", "enabled"))
%>