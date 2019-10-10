<script language="jscript" runat="server">
	function Format(str) {
		if (arguments.length===1) return str;
		var args=arguments;
		var ret=str.replace(/\{(\d+)\}/g,
			function(s, num) {
				return String(args[parseInt(num)+1]);
			}
		);
		return ret;
	}
</script><%
Function toSqlString(ByVal str)
	If IsNull(str) Then
		toSqlString="NULL"
		Exit Function
	End If
	str=Replace(Replace(Replace(Replace(str,"'","''"),vbCrLf,"'+CHAR(13)+CHAR(10)+'"),vbCr,"'+CHAR(13)+'"),vbLf,"'+CHAR(10)+'")
	If Len(str)=0 Then
		str="NULL"
	Else
		str="N'"&str&"'"
	End If
	toSqlString=str
End Function

Function toSqlNumber(ByVal str)
	If IsNull(str) Then
		toSqlNumber="NULL"
		Exit Function
	End If
	toSqlNumber=str
End Function

Function toPlainString(ByVal str)
	If IsNull(str) Then Exit Function
	str=Replace(str,"""","&quot;")
	str=Replace(str,"<","&lt;")
	str=Replace(str,">","&gt;")
	str=Replace(str,vbNewLine,"<br/>")
	str=Replace(str,vbCr,"<br/>")
	str=Replace(str,vbLf,"<br/>")
	toPlainString=str
End Function

Function toJsString(ByVal str)
	If IsNull(str) Then Exit Function
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
	toJsString = Join(haystack, "")
End Function

Function toPlainText(ByVal str)
	Dim r,arr,match
	Set r=New RegExp
	r.Pattern="<[^>]*>|&[A-z]+;"
	r.IgnoreCase=True
	r.Global=True
	Set arr=r.Execute(str)
	For Each match In arr
		str=Replace(str,match.Value,vbNullString)
	Next
	toPlainText=str
End Function

Function toEscapeString(ByVal str)
	Dim i,high,low,char,ret
	For i=1 To Len(str)
		char=Hex(AscW(Mid(str,i,1)))
		char=String(4-Len(char),"0")&char
		low=Right(char,2)
		high=Left(char,2)
		ret=ret&"%"&low&"%"&high
	Next
	toEscapeString=ret
End Function

Function toNumericString(ByVal str)
	str=Trim(str)
	If IsNull(str) Or IsEmpty(str) Then
		toNumericString="0"
		Exit Function
	ElseIf Not IsNumeric(str) Then
		toNumericString="0"
		Exit Function
	End If
	If Left(str,1)="." Then
		str="0"&str
	End If
	toNumericString=str
End Function

Function toUnsignedInt(ByVal str)
	str=Trim(str)
	If IsNull(str) Or IsEmpty(str) Then
		toUnsignedInt=-1
		Exit Function
	ElseIf Not IsNumeric(str) Then
		toUnsignedInt=-1
		Exit Function
	End If
	toUnsignedInt=Int(str)
End Function

Function toDateTime(d,fmt)
	If IsNull(d) Then
		toDateTime="0"
		Exit Function
	End If
	If d="0" Then
		toDateTime="未定"
	Else
		toDateTime=FormatDateTime(d,fmt)
	End If
End Function

Function toYearMonthDate(ByVal year,ByVal month,ByVal date)
	' 返回形如 yyyy年m月d日 的日期格式
	'If Len(month)=1 Then month="0"&month
	'If Len(date)=1 Then date="0"&date
	toYearMonthDate=year&"年"&month&"月"&date&"日"
End Function

Function toYearMonth(ByVal year,ByVal month)
	' 返回形如 yyyy年m月 的日期格式
	'If Len(month)=1 Then month="0"&month
	toYearMonth=year&"年"&month&"月"
End Function

Function toDataSizeString(value)
	Dim unit:unit=Array("B","KiB","MiB","GiB")
	Dim d,i:i=0
	d=value
	Do While d>=1024
		d=d/1024
		i=i+1
	Loop
	toDataSizeString=Round(d,2)&" "&unit(i)
End Function

Function toSnakeCase(ByVal str)
	If IsNull(str) Then Exit Function
	Dim lastIndex:lastIndex=1
	Dim ret
	Dim re:Set re=New RegExp
	re.Pattern="([a-z])([A-Z])"
	re.Global=True
	Set matches=re.Execute(str)
	For Each match In matches
		Dim value: value=match.SubMatches(0)&"_"&LCase(match.SubMatches(1))
		ret=ret&LCase(Mid(str,lastIndex,match.FirstIndex-lastIndex+1))&value
		lastIndex=match.FirstIndex+match.Length+1
	Next
	ret=ret&LCase(Mid(str,lastIndex))
	Set re=Nothing
	toSnakeCase=ret
End Function

Function ArrayJoin(arr,delim)
	If Not IsArray(arr) Then
		ArrayJoin=arr
		Exit Function
	End If
	Dim i,ret
	For i=0 To UBound(arr)
		If i>0 Then ret=ret&delim
		ret=ret&arr(i)
	Next
	ArrayJoin=ret
End Function

Function isMatched(pattern,s,ignoreCase)
	' 判断指定字符串是否满足指定模式
	Dim regEx:Set regEx=New RegExp
	regEx.Pattern=pattern
	regEx.IgnoreCase=ignoreCase
	isMatched=regEx.Test(s)
	Set regEx=Nothing
End Function

Function generatePassword()
	' 生成随机密码序列
	Dim i,j,lenPwdChar,lenPwd,ret
	Const pwdchar="Aa0Bb1Cc2Dd3Ee4Ff5Gg6Hh7Ii8Jj9Kk9L8Mm7Nn6Oo5Pp4Qq3Rr2Ss1Tt0UuVvWwXxYyZzAa0Bb1Cc"
	lenPwdChar=Len(pwdchar)
	lenPwd=7
	Randomize()
	For i=1 To lenPwd
		j=Int(Rnd()*lenPwdChar)
		ret=ret&Mid(pwdchar,j,1)
	Next
	generatePassword=ret
End Function
%>