<%
' JSON 序列输出类
Class JSONWriter

    Public Default Function output(res, data)
        res.AddHeader "Content-Type", "application/json"
        If res.Buffer Then res.Clear()
        res.Write stringify(data)
    End Function

    Function parse(str) 	' TODO: parse 的完整实现
        Dim obj
        Eval("obj="&str)
        parse = obj
    End Function
    
    Function stringify(obj)
        Dim obj_type: obj_type=TypeName(obj)
        If obj_type="String" Or obj_type="Date" Then
            stringify=""""&jsEncode(obj)&"""": Exit Function
        ElseIf obj_type="Empty" Then
            stringify="undefined": Exit Function
        ElseIf InStr("Boolean|Integer|Long|Double|Byte|Currency|Single", obj_type) Then
            stringify=LCase(CStr(obj)): Exit Function
        ElseIf IsNull(obj) Then
            stringify="null": Exit Function
        ElseIf IsArray(obj) Then
            Dim arr_prop(): ReDim arr_prop(UBound(obj))
            Dim j
            For j=0 To UBound(obj)
                arr_prop(j)=stringify(obj(j))
            Next
            stringify="["&Join(arr_prop,",")&"]": Exit Function
        ElseIf obj_type<>"Dictionary" Then
            stringify="""[Object "&obj_type&"]""": Exit Function
        End If
        Dim arr: ReDim arr(obj.Count-1)
        Dim prop, i: i=0
        For Each prop In obj.Keys()
            If Not IsEmpty(obj.Item(prop)) Then
                Dim value
                If IsObject(obj.Item(prop)) Then
                    Set value=obj.Item(prop)
                Else
                    value=obj.Item(prop)
                End If
                Dim prop_snake: prop_snake=jsEncode(toSnakeCase(prop))
                arr(i)=Format("""{0}"":{1}", prop_snake, stringify(value))
            Else
                arr(i)=Format("""{0}"":undefined", prop_snake)
            End If
            i=i+1
        Next
        stringify="{"&Join(arr,",")&"}"
    End Function

	Function jsEncode(str)
		Dim charmap(127), haystack()
		charmap(8)  = "\b"
		charmap(9)  = "\t"
		charmap(10) = "\n"
		charmap(12) = "\f"
		charmap(13) = "\r"
		charmap(34) = "\"""
		charmap(47) = "\/"
		charmap(92) = "\\"

		Dim strlen : strlen = Len(str) - 1
		ReDim haystack(strlen)

		Dim i, charcode
		For i = 0 To strlen
			haystack(i) = Mid(str, i + 1, 1)

			charcode = AscW(haystack(i)) And 65535
			If charcode < 127 Then
				If Not IsEmpty(charmap(charcode)) Then
					haystack(i) = charmap(charcode)
				ElseIf charcode < 32 Then
					haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
				End If
			Else
				haystack(i) = "\u" & Right("000" & Hex(charcode), 4)
			End If
		Next

		jsEncode = Join(haystack, "")
    End Function
    
End Class
%>