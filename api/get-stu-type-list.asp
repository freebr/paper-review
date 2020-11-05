<%
'==========================================
' API 名称：    get-stu-type-list
' API 功能：    提供现有学生类型的列表
' API 输出类型：JSON
' 修订日期：    2019-05-02
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main()
    Dim data: Set data=CreateDictionary()
    Dim conn,rs,sql,count
    ConnectDb conn
    sql="SELECT TEACHTYPE_ID id, TEACHTYPE_NAME name FROM ViewStudentTypeInfo WHERE (?&StuTypeBitwise)<>0"
    Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("ManageStuTypes",adInteger,4,Session("AdminType")("ManageStuTypes")))
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