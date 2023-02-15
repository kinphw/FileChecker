Attribute VB_Name = "ClearFormula_Dist"
Sub ClearFormula(control As IRibbonControl)

Application.ScreenUpdating = False

On Error Resume Next

Dim i As Integer
i = 0
Dim arr() As Variant
ReDim arr(0) As Variant

For Each sheet In Application.ActiveWorkbook.Sheets

    sheet.Activate
    ActiveSheet.Cells.SpecialCells(xlCellTypeFormulas, 16).ClearContents
    
    If Err.Number <> 0 Then
        Err.Clear
    Else
        'MsgBox Sheet.Name + "��Ʈ ���� ����"
        arr(i) = sheet.Name
        i = i + 1
        ReDim Preserve arr(0 To i)
    End If
    
Next

Application.ScreenUpdating = True

Dim str As String

If i <> 0 Then
    For Each Item In arr
        str = str + Item + vbCrLf
    Next Item
    str = str + "�� ��Ʈ�� �������� �Ϸ�"

Else
    str = ActiveWorkbook.Name + "���Ͽ� ������ �����ϴ�."

End If

MsgBox str

End Sub
