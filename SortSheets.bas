Attribute VB_Name = "SortSheets"
Sub SortSheets()

Dim i As Integer
Dim j As Integer
Dim iAnswer As VbMsgBoxResult

iAnswer = MsgBox("��Ʈ�� ���� �������� �����Ͻðڽ��ϱ�?" & Chr(10) _
& "���� ������ �ƴϿ��� �����ּ���", _
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

