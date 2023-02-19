Attribute VB_Name = "FilterHere"

Sub FilterNot0()

'0이 아닌 것 필터

Dim rng As Range
Set rng = Selection

Dim rngP As Range
Set rngP = Selection.EntireRow

'rngP.AutoFilter(Field:=rng.Column,Criteria1:="<>" & 0)

Call rngP.AutoFilter(Field:=rng.Column, Criteria1:="<>" & 0)

End Sub


'Sub Auto_Open()

'Application.OnKey "^l", FilterNot0

'End Sub
