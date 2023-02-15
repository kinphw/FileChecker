Attribute VB_Name = "IsThere_Dist"
Function IsThere(a As Range, b As Range)

'인수 a: VLookup 대상
'인수 b: VLookup 범위

'a = ActiveSheet.Range("A1")
'b = ActiveSheet.Range("B1:B2")

'중간변수
Dim v1 As Variant
Dim v2 As Variant

'1. Vlookup 실시 결과값 = b

v1 = Application.VLookup(a, b, 1, 0) '왜인지는 알수 없지만 이거로
v2 = Not WorksheetFunction.IsError(v1)

IsThere = v2
'
'MsgBox v2

End Function
