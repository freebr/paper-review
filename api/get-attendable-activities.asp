<%
'==========================================
' API 名称：    get-attendable-activities
' API 功能：    提供学生可参加的评阅活动信息
' API 输出类型：JSON
' 修订日期：    2019-05-12
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../student/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main()
    Dim data: Set data=CreateDictionary()
    Dim conn,rs,sql,count
    Connect conn
    sql="SELECT Id,Name,SemesterId,SemesterName,StuType,StuTypeName FROM ViewActivities WHERE Valid=1 AND IsOpen=1 AND (?&StuType)<>0"
    Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("CurrentStuType",adInteger,4,2^(Session("StuType")-1)))
    Set rs=ret("rs")
    count=ret("count")
    Dim arr(): ReDim arr(count-1)
    Dim i: i=0
    Do While Not rs.EOF
        Dim dictItem: Set dictItem=CreateDictionary()
        Dim j
        For j=0 To rs.Fields.Count-1
            dictItem.Add rs.Fields(j).Name, rs(j).Value
        Next
        Set arr(i)=dictItem
        i=i+1
        rs.MoveNext()
    Loop
    data.Add "status", "ok"
    data.Add "data", arr

    CloseRs rs
    CloseConn conn
    
    Call (new JSONWriter)(Response, data)
End Function

Call main()
%>