Attribute VB_Name = "FilterHere"

Sub FilterNot0()

'0�� �ƴ� �� ����

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
