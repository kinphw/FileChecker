Attribute VB_Name = "Help_Dist"
Sub Help(control As IRibbonControl)

Dim str As String

str = "제작자 : 박병"
str = str + vbCrLf
str = str + "v2 DD 230215"
str = str + vbCrLf
str = str + "기능 : 메모(노트) 일괄삭제/오류함수 일괄삭제"
str = str + vbCrLf
str = str + "함수 : IsThere() => 영역에 존재하는지 여부를 True/False로 반환"

MsgBox str

End Sub
