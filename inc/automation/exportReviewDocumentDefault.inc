<%
Const msoTextBox=17
Const wdExportFormatPDF=17
Const wdGoToPage=1
Const wdGoToAbsolute=1
Const wdStatisticPages=2
Function exportReviewDocumentDefault(doc,stu_type,filepath,filepath2,filepath3)
    Dim range,range2,text
    Dim i,l,font_size
    Dim val,bValid
    For i=1 To doc.Shapes.Count
        If doc.Shapes(i).Type=msoTextBox Then
            bValid=True
            Set range=doc.Shapes(i).TextFrame.TextRange
            text=Trim(Left(range.Text,Len(range.Text)-1))
            Select Case text
            Case "$expert_name":val=info("ExpertName")
            Case "$expert_pro_duty":val=info("ExpertProDuty")
            Case "$expert_expertise":val=info("ExpertExpertise")
            Case "$expert_workplace":val=info("ExpertWorkplace")
            Case "$expert_address":val=info("ExpertAddress")
            Case "$expert_mailcode":val=info("ExpertMailcode")
            Case "$expert_tel1":val=info("ExpertTel1")
            Case "$expert_tel2":val=info("ExpertTel2")
            Case "$expert_master_level":val=String(info("ExpertMasterLevel"),vbTab)&"√"
            Case "$correlation_level":val=String(info("CorrelationLevel"),vbTab)&"√"
            Case "$review_result"
                If stu_type=5 Then
                    val=String((info("ReviewResult")-1)\2,vbNewLine)&String((info("ReviewResult")-1)Mod 2+1,vbTab)&"√"
                Else
                    val=String(info("ReviewResult"),vbTab)&"√"
                End If
            Case "$review_level":val=String(info("ReviewLevel"),vbTab)&"√"
            Case Else:bValid=False
            End Select
            If bValid Then
                range.Select()
                range.Find.Execute text,True,,,,,,1,,toWordString(val),True
            End If
        End If
    Next
    Set range=doc.Range()
    range.Select()
    Do
        bFind=range.Find.Execute("$author",True,,,,,,1,,toWordString(info("Author")),True)
    Loop While bFind
    
    ' 调整字体大小
    l=Len(info("Subject"))
    If l<=17 Then
        font_size=16
    ElseIf l<=20 Then
        font_size=14
    ElseIf l<=23 Then
        font_size=12
    Else
        font_size=10.5
    End If
    Set range2=doc.Tables(1).Range
    If range2.Find.Execute("$thesis_subject",True) Then
        range2.Font.Size=font_size
    End If
    Do
        bFind=range.Find.Execute("$thesis_subject",True,,,,,,1,,toWordString(info("Subject")),True)
    Loop While bFind
    range.Find.Execute "$tutorinfo",True,,,,,,1,,toWordString(info("TutorInfo")),True
    range.Find.Execute "$spec",True,,,,,,1,,toWordString(info("Speciality")),True
    range.Find.Execute "$researchway",True,,,,,,1,,toWordString(info("Researchway")),True
    range.Find.Execute "$submit_review_date",True,,,,,,1,,toWordString(info("SubmitReviewDate")),True
    If range.Find.Execute("$eval_text",True,,,,,,1) Then
        range.Text=info("Comment")
    End If
    If Len(info("Scores")) Then
        Dim scores:scores=Split(info("Scores"),",")
        ' 对ME评阅书进行处理
        i=doc.ComputeStatistics(wdStatisticPages)
        j=doc.Tables(3).Range.End
        Set r1=doc.GoTo(wdGoToPage,wdGoToAbsolute,i-1)
        Set r2=doc.GoTo(wdGoToPage,wdGoToAbsolute,i)
        If Len(doc.Range(r1.Start,r2.Start))=2 Then
            ' 删除多余的换页符
            doc.Range(j,j).Delete()
        End If
        Set r1=Nothing
        Set r2=Nothing
        For i=0 To UBound(scores)
            range.Find.Execute "$score"&(i+1),True,,True,,,,1,,toWordString(Int(scores(i))),True
        Next
    End If
    If Len(info("ScoreParts")) Then
        Dim scoreparts:scoreparts=Split(info("ScoreParts"),",")
        For i=0 To UBound(scoreparts)
            range.Find.Execute "$scorep"&(i+1),True,,True,,,,1,,toWordString(Int(scoreparts(i))),True
        Next
    End If
    range.Find.Execute "$total_score",True,,True,,,,1,,toWordString(info("TotalScore")),True
    ' 生成完整版评阅书
    doc.ExportAsFixedFormat filepath,wdExportFormatPDF,,1
    ' 生成不含学生信息的评阅书
    doc.Tables(1).Delete()
    If reviewfile_type=1 Then
        doc.Range.Paragraphs(4).Range.Delete()
    Else
        doc.Range.Paragraphs(3).Range.Delete()
    End If
    doc.Tables(2).Cell(1,3).Delete()
    doc.Tables(2).Cell(1,2).Delete()
    doc.Tables(2).Cell(1,1).Range.Text="论文题目"
    doc.Tables(2).Cell(1,2).SetWidth 412.65,1
    doc.ExportAsFixedFormat filepath2,wdExportFormatPDF,,1
    ' 生成不含专家信息的评阅书
    doc.Undo 6
    doc.Tables(2).Range.Delete()
    doc.ExportAsFixedFormat filepath3,wdExportFormatPDF,,1
    Set range2=Nothing
    Set range=Nothing
    exportReviewDocumentDefault=1
End Function
%>