<%
'==========================================
' API 名称：    get-review-records
' API 功能：    提供指定论文的历史评阅记录信息
' API 输出类型：JSON
' 修订日期：    2019-05-19
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
    sql="SELECT Id,ReviewOrder,ReviewOrderText,ReviewerId,ReviewerName,ReviewTime,OverallRatingText,DefenceOpinionText,ReviewFile,Comment,DisplayStatus,DisplayStatusModifiedByName,Creator,CreatorName FROM ViewReviewRecords WHERE DissertationId=?"
    Dim ret:Set ret=ExecQuery(conn,sql,CmdParam("DissertationId",adInteger,4,arg("id")))
    Set rs=ret("rs")
    count=ret("count")
    Dim arr()
	If count=0 Then
        Dim dict:Set dict = CreateDictionary()
		dict.Add "Id","0"
        ReDim arr(0):Set arr(0) = dict
	Else
        ReDim arr(count-1)
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
    End If
    data.Add "status", "ok"
    data.Add "data", arr

    CloseRs rs
    CloseConn conn
    
    Call (new JSONWriter)(Response, data)
End Function

Call main("id")
%>