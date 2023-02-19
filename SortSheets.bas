Attribute VB_Name = "SortSheets"
Sub SortSheets()

Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult

iAnswer = MsgBox("시트를 오름 차순으로 정렬하시겠습니까?" & Chr(10) _
& "내림 차순은 아니오를 눌러주세요", _
vbYesNoCancel + vbQuestion + vbDefaultButton1, "Sort Worksheets")
    For i = 1 To Sheets.Count
        For j = 1 To Sheets.Count - 1
            If iAnswer = vbYes Then
                If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                    Sheets(j).Move After:=Sheets(j + 1)
                End If
            ElseIf iAnswer = vbNo Then
                If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then Sheets(j).Move After:=Sheets(j + 1)
            End If
        Next j
    Next i
End Sub

