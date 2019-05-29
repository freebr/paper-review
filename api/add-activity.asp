<%
'==========================================
' API 名称：    add-activity
' API 功能：    新增评阅活动
' API 输入类型：POST
' API 输出类型：JSON
' 修订日期：    2019-05-02
'==========================================
%><!--#include file="../inc/global.inc"-->
<!--#include file="../admin/common.asp"-->
<!--#include file="../inc/api.asp"--><%

Function main(args)
    Dim data: Set data=CreateDictionary()
    Dim arg: Set arg=CreateDictionary()
    ensureArgument args, arg, data
    Dim stu_type:stu_type=0
    Dim arrStuTypes:arrStuTypes=Split(arg("stu_type"), ",")
    For i=0 To UBound(arrStuTypes)
        stu_type=stu_type+2^(Int(arrStuTypes(i))-1)
    Next
    Dim conn,sql,sql2,count
    Connect conn
    sql="INSERT INTO Activities (Name,SemesterId,StuType,IsOpen,CreatedAt,Creator) VALUES (?,?,?,?,?,?)"
    On Error Resume Next
    count=ExecNonQuery(conn,sql,_
        CmdParam("Name",adVarWChar,50,arg("name")),_
        CmdParam("SemesterId",adInteger,4,arg("semester")),_
        CmdParam("StuType",adInteger,4,stu_type),_
        CmdParam("IsOpen",adBoolean,1,arg("is_open")),_
        CmdParam("CreatedAt",adDate,4,Now),_
        CmdParam("Creator",adInteger,4,Session("id")))
    
    Dim ret:Set ret=ExecQuery(conn,"SELECT CAST(@@IDENTITY AS int)")
    Dim rs:Set rs=ret("rs")
    count=ret("count")
    Dim new_id:new_id=rs(0).Value
    CloseRs rs

    ' 获取需记录开放时间的所有环节Id
    Const count_client_type=3
    Dim i,j,k
    Dim sectionIds: ReDim sectionIds(count_client_type)
    For i=1 To count_client_type
        sql="SELECT Id FROM Sections WHERE ClientType=?"
        Set ret=ExecQuery(conn,sql,CmdParam("ClientType",adInteger,4,i))
        Set rs=ret("rs")
        count=ret("count")
        Dim arr: ReDim arr(count)
        For j=1 To count
            arr(j)=rs(0).Value
            rs.MoveNext()
        Next
        sectionIds(i)=arr
    Next

    ' 新增所有环节的开放时间记录
    Dim params(4): Set params(0)=CmdParam("ActivityId",adInteger,4,new_id)
    Dim mail_templates: Set mail_templates=CreateDictionary()
    With mail_templates
        .Add "lwsstzyj(xs)", "论文送审通知邮件（学生）"
        .Add "lwsstzyj(ds)", "论文送审通知邮件（导师）"
        .Add "lwdpytzyj", "论文待评阅通知邮件"
        .Add "lwdpytzdx", "论文待评阅通知短信"
        .Add "lwshtzyj", "论文审核通知邮件"
        .Add "lwshwtgtzyj", "论文审核未通过通知邮件"
        .Add "lwshtgtzyj", "论文审核通过通知邮件"
        .Add "pyyjqrtzyj", "评阅意见确认通知邮件"
        .Add "xxdrtzyj(xs)", "信息导入通知邮件（学生）"
        .Add "xxdrtzyj(ds)", "信息导入通知邮件（导师）"
        .Add "dbsxtzyj", "待办事项通知邮件"
    End With
    Dim mail_template_keys: mail_template_keys=mail_templates.Keys()
    Dim new_time: new_time=Now
    sql="INSERT INTO ActivityPeriods (ActivityId,StuType,SectionId,StartTime,EndTime) VALUES (?,?,?,?,?)"
    sql2="INSERT INTO ActivityMailTemplates (ActivityId,StuType,Name,MailSubject,MailContent) VALUES (?,?,?,?,?)"
    For i=0 To UBound(arrStuTypes)
        Set params(1)=CmdParam("StuType",adInteger,4,arrStuTypes(i))
        For j=1 To UBound(sectionIds)
            For k=1 To UBound(sectionIds(j))
                Set params(2)=CmdParam("SectionId",adInteger,4,sectionIds(j)(k))
                Set params(3)=CmdParam("StartTime",adInteger,4,new_time)
                Set params(4)=CmdParam("EndTime",adInteger,4,new_time)
                count=ExecNonQuery(conn,sql,params(0),params(1),params(2),params(3),params(4))
            Next
        Next
        ' 新增通知邮件（短信）模板
        For j=0 To UBound(mail_template_keys)
            Set params(2)=CmdParam("Name",adVarWChar,50,mail_template_keys(j))
            Set params(3)=CmdParam("MailSubject",adVarWChar,50,mail_templates(mail_template_keys(j)))
            Set params(4)=CmdParam("MailContent",adVarWChar,3000,"")
            count=ExecNonQuery(conn,sql2,params(0),params(1),params(2),params(3),params(4))
        Next
    Next

    If Err.Number Then
        data.Add "status", "error"
        data.Add "msg", Err.Description
    Else
        data.Add "status", "ok"
        data.Add "id", new_id
    End If
    On Error GoTo 0
    CloseRs rs
    CloseConn conn

    Call (new JSONWriter)(Response, data)
End Function

Call main(Array("name", "semester", "stu_type", "is_open"))
%>