<%
'=========================================
' 专业学位论文电子评阅系统全局成员定义文件
' 2016-09-11
'=========================================

' TASK_PROGRESS
Const tpNone=0
Const tpTbl1Uploaded=1
Const tpTbl1Unpassed=2
Const tpTbl1Passed=3
Const tpTbl2Uploaded=4
Const tpTbl2Unpassed=5
Const tpTbl2Passed=6
Const tpTbl3Uploaded=7
Const tpTbl3Unpassed=8
Const tpTbl3Passed=9
Const tpTbl4Uploaded=10
Const tpTbl4Unpassed=11
Const tpTbl4Passed=12
' REVIEW_STATUS
Const rsNone=0
Const rsDetectThesisUploaded=1
Const rsNotAgreeDetect=2
Const rsAgreeDetect=3
Const rsDetectUnpassed=4
Const rsRedetectUnpassed=5
Const rsRedetectPassed=6
Const rsNotAgreeReview=7
Const rsAgreeReview=8
Const rsMatchExpert=9
Const rsReviewed=10
Const rsReviewEval=11
Const rsModifyThesisUploaded=12
Const rsModifyUnpassed=13
Const rsModifyPassed=14
Const rsDefenceEval=15
Const rsInstructEval=16
Const rsFinalThesisUploaded=17

Const STUTYPE_COUNT=4
Const OPRTYPE_COUNT=8

Class StudentClientInfo
	Private opentime_start(),opentime_end()
	Private m_client_status()
	Private m_systemstatus
	Private Sub Class_Initialize()
		ReDim opentime_start(OPRTYPE_COUNT)
		ReDim opentime_end(OPRTYPE_COUNT)
		ReDim m_client_status(STUTYPE_COUNT,OPRTYPE_COUNT)
	End Sub
	Public Function getOpentime(Opr,TimeType)
		If TimeType=STUCLI_OPENTIME_START Then
			getOpentime=opentime_start(Opr)
		ElseIf TimeType=STUCLI_OPENTIME_END Then
			getOpentime=opentime_end(Opr)
		End If
	End Function
	Public Function setOpentime(Opr,TimeType,ByVal Value)
		If IsNull(Value) Then Value=0
		If TimeType=STUCLI_OPENTIME_START Then
			opentime_start(Opr)=Value
		ElseIf TimeType=STUCLI_OPENTIME_END Then
			opentime_end(Opr)=Value
		End If
		setOpentime=1
	End Function
	Public Function getClientStatus(StuType,Opr)
		Dim i
		Select Case StuType
		Case STUCLI_STUTYPE_ME:i=1
		Case STUCLI_STUTYPE_MBA:i=2
		Case STUCLI_STUTYPE_EMBA:i=3
		Case STUCLI_STUTYPE_MPACC:i=4
		Case Else
			getClientStatus=0
			Exit Function
		End Select
		getClientStatus=m_client_status(i,Opr)
	End Function
	Public Function setClientStatus(val)
		Dim arr,i,j
		arr=Split(val,",")
		For i=1 To STUTYPE_COUNT
			For j=1 To OPRTYPE_COUNT
				m_client_status(i,j)=arr(OPRTYPE_COUNT*i+j-OPRTYPE_COUNT)
			Next
		Next
	End Function
	Public Function isOpenFor(StuType,Opr)
		Dim i
		Select Case StuType
		Case STUCLI_STUTYPE_ME:i=1
		Case STUCLI_STUTYPE_MBA:i=2
		Case STUCLI_STUTYPE_EMBA:i=3
		Case STUCLI_STUTYPE_MPACC:i=4
		Case Else
			isOpenFor=False
			Exit Function
		End Select
		isOpenFor=m_client_status(i,Opr)="1" And DateDiff("d",opentime_start(Opr),Now)>=0 And DateDiff("d",opentime_end(Opr),Now)<=0
	End Function
	Public Property Get SystemStatus()
		SystemStatus=m_systemstatus
	End Property
	Public Property Let SystemStatus(val)
		m_systemstatus=val
	End Property
End Class

Function hasPrivilege(pArr,privilege)
	If IsEmpty(pArr) Then 
		hasPrivilege=False
		Exit Function
	End If
	pArr=split(pArr,",")
	For i=0 To Ubound(pArr)-1
		If pArr(i)=privilege Then
      hasPrivilege=True
      Exit Function
	  End If
	Next
	hasPrivilege=False
End Function

Const STUCLI_OPR_TABLE1 = 1
Const STUCLI_OPR_TABLE2 = 2
Const STUCLI_OPR_TABLE3 = 3
Const STUCLI_OPR_TABLE4 = 4
Const STUCLI_OPR_DETECT_REVIEW = 5
Const STUCLI_OPR_REVIEW = 6
Const STUCLI_OPR_MODIFY = 7
Const STUCLI_OPR_FINAL = 8
Const STUCLI_OPENTIME_START = 1
Const STUCLI_OPENTIME_END = 2
Const STUCLI_STUTYPE_ME = 5
Const STUCLI_STUTYPE_MBA = 6
Const STUCLI_STUTYPE_EMBA = 7
Const STUCLI_STUTYPE_MPACC = 9
Const STUCLI_STATUS_CLOSED = 0
Const STUCLI_STATUS_OPEN = 1
Const STUCLI_STATUS_DEBUG = 2

Dim arrStuTypeId:arrStuTypeId=Array(0,5,6,7,9)
Dim arrStuTypeName:arrStuTypeName=Array("","ME","MBA","EMBA","MPAcc")
%>