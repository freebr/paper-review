Sub extractMBAInfo()
'
' extractMBAInfo 宏
'
'
    Dim doc As Document: Set doc = ActiveDocument
    Dim ret As String, tmp As String
    ' 学号
    ret = ret & compactString(doc.Tables(2).Cell(7, 2).Range.text, True) & ","
    ' 评语
    Dim rg As Range: Set rg = doc.Tables(12).Cell(2, 1).Range
    Dim i
    For i = 2 To rg.Sentences.Count
        tmp = tmp & rg.Sentences(i).text
    Next
    ret = ret & compactString(tmp, False)
    
    ' 评分
    Dim tbl As Table: Set tbl = doc.Tables(10)
    For i = 2 To tbl.Rows.Count
        ret = ret & ","
        If i = tbl.Rows.Count Then  ' 加权总分
            ret = ret & getScore(tbl.Cell(i, 2).Range.text)
        Else
            If tbl.Cell(i, 3).Next.RowIndex = i Then '一级指标得分
                ret = ret & getScore(tbl.Cell(i, 4).Range.text) & ","
            End If
            ret = ret & getScore(tbl.Cell(i, 3).Range.text)
        End If
    Next
    
    ' 评阅专家信息
    ' 专业技术职务
    ret = ret & "," & compactString(doc.Tables(4).Cell(1, 4).Range.text, True)
    ' 学科专长
    ret = ret & "," & compactString(doc.Tables(5).Cell(1, 2).Range.text, True)
    ' 单位名称（含院系）
    ret = ret & "," & compactString(doc.Tables(6).Cell(1, 2).Range.text, True)
    ' 通信地址
    ret = ret & "," & compactString(doc.Tables(7).Cell(1, 2).Range.text, True)
    ' 联系电话
    ret = ret & "," & compactString(doc.Tables(7).Cell(1, 4).Range.text, True)
    
    ' 对论文涉及内容的熟悉程度
    Dim arrMasterLevelText: arrMasterLevelText = Array("", "很熟悉(√)", "熟悉(√)", "一般(√)")
    For i = 1 To UBound(arrMasterLevelText)
        If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 对学位论文的总体评价
    For i = 1 To 4
        If doc.Tables(13).Cell(2, i + 1).Range.Find.Execute("√") Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 是否同意申请论文答辩
    Set rg = doc.Tables(13).Cell(3, 2).Range
    Dim j As Integer: j = 1
    For i = 1 To rg.Paragraphs.Count
        tmp = compactString(rg.Paragraphs(i).Range.text, True)
        If Len(tmp) > 1 Then
            If rg.Paragraphs(i).Range.Find.Execute("√") Then
                ret = ret & "," & j
                Exit For
            End If
            j = j + 1
        End If
    Next
    
    Debug.Print ret
    
End Sub

Sub extractMEMInfo()
'
' extractMEMInfo 宏
'
'
    Dim doc As Document: Set doc = ActiveDocument
    Dim ret As String
    
    ' 学号
    ret = ret & compactString(doc.Tables(9).Cell(1, 4).Range.Text, True)
    ' 评语
    Dim rg As Range: Set rg = doc.Tables(9).Cell(3, 1).Range
    Dim i
    For i = 2 To rg.Sentences.Count
        tmp = tmp & rg.Sentences(i).Text
    Next
    ret = ret & "," & compactString(tmp, False)
    
    ' 评分
    Dim tbl As Table: Set tbl = doc.Tables(10)
    For i = 2 To tbl.Rows.Count
        ret = ret & ","
        If i = tbl.Rows.Count Then  ' 加权总分
            ret = ret & getScore(tbl.Cell(i, 2).Range.Text)
        Else
            ret = ret & getScore(tbl.Cell(i, 5).Range.Text)
        End If
    Next
    
    ' 评阅专家信息
    ' 专业技术职务
    ret = ret & "," & compactString(doc.Tables(4).Cell(1, 4).Range.Text, True)
    ' 学科专长
    ret = ret & "," & compactString(doc.Tables(5).Cell(1, 2).Range.Text, True)
    ' 单位名称（含院系）
    ret = ret & "," & compactString(doc.Tables(6).Cell(1, 2).Range.Text, True)
    ' 通信地址
    ret = ret & "," & compactString(doc.Tables(7).Cell(1, 2).Range.Text, True)
    ' 联系电话
    ret = ret & "," & compactString(doc.Tables(7).Cell(1, 4).Range.Text, True)
    
    ' 对论文涉及内容的熟悉程度
    Dim arrMasterLevelText: arrMasterLevelText = Array(, "很熟悉(√)", "熟悉(√)", "一般(√)")
    For i = 1 To UBound(arrMasterLevelText)
        If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 对学位论文的总体评价
    For i = 1 To 4
        If doc.Tables(11).Cell(2, i + 1).Range.Find.Execute("√") Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 是否同意申请论文答辩
    Set rg = doc.Tables(11).Cell(3, 2).Range
    Dim j As Integer: j = 1
    For i = 1 To rg.Paragraphs.Count
        tmp = compactString(rg.Paragraphs(i).Range.Text, True)
        If Len(tmp) > 1 Then
            If rg.Paragraphs(i).Range.Find.Execute("√") Then
                ret = ret & "," & j
                Exit For
            End If
            j = j + 1
        End If
    Next
    
    Debug.Print ret
    
End Sub

Sub extractMPACCInfo()
'
' extractMPACCInfo 宏
'
'
    Dim doc As Document: Set doc = ActiveDocument
   Dim ret As String, tmp As String
    ' 学号
    ret = ret & compactString(doc.Tables(9).Cell(1, 4).Range.Text, True)
    ' 评语
    Dim rg As Range: Set rg = doc.Tables(9).Cell(2, 1).Range
    Dim i
    For i = 2 To rg.Sentences.Count
        tmp = tmp & rg.Sentences(i).Text
    Next
    ret = ret & "," & compactString(tmp, False)
    
    ' 评阅专家信息
    ' 专业技术职务
    ret = ret & "," & compactString(doc.Tables(4).Cell(1, 4).Range.Text, True)
    ' 学科专长
    ret = ret & "," & compactString(doc.Tables(5).Cell(1, 2).Range.Text, True)
    ' 单位名称（含院系）
     ret = ret & "," & compactString(doc.Tables(6).Cell(1, 2).Range.Text, True)
    ' 通信地址
     ret = ret & "," & compactString(doc.Tables(7).Cell(1, 2).Range.Text, True)
    ' 联系电话
     ret = ret & "," & compactString(doc.Tables(7).Cell(1, 4).Range.Text, True)
    
    ' 对论文涉及内容的熟悉程度
    Dim arrMasterLevelText: arrMasterLevelText = Array("", "很熟悉(√)", "熟悉(√)", "一般(√)")
    For i = 1 To UBound(arrMasterLevelText)
        If doc.Tables(7).Cell(2, 1).Range.Find.Execute(arrMasterLevelText(i)) Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 对学位论文的总体评价
    For i = 1 To 4
        If doc.Tables(10).Cell(2, i + 1).Range.Find.Execute("√") Then
            ret = ret & "," & i
            Exit For
        End If
    Next
    
    ' 是否同意申请论文答辩
    Set rg = doc.Tables(10).Cell(3, 2).Range
    Dim j: j = 1
    For i = 1 To rg.Paragraphs.Count
        tmp = compactString(rg.Paragraphs(i).Range.Text, True)
        If Len(tmp) > 1 Then
            If rg.Paragraphs(i).Range.Find.Execute("√") Then
                ret = ret & "," & j
                Exit For
            End If
            j = j + 1
        End If
    Next
    
    Debug.Print ret
    
End Sub

Function compactString(ByVal str As String, singleLine As Boolean)
    str = Replace(Replace(str, Chr(7), ""), ChrW(160), "")
    If singleLine Then
        str = Replace(Replace(Replace(str, vbCr, ""), vbLf, ""), Chr(11), "")
    Else
        str = Replace(str, Chr(11), vbNewLine)
    End If
    compactString = str
End Function

Function getScore(ByVal str As String)
    getScore = Round(Replace(compactString(str, True), "分", ""))
End Function