Attribute VB_Name = "KillName"
Sub ����()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual


On Error Resume Next

'For Each c In ThisWorkbook.Names
For Each c In ActiveWorkbook.Names 'ActiveWorkbook���� �ؾ� �۵��� (��⼳���� �߰����.xlam�� �ƴϴϱ�)
    c.Delete
Next c


MsgBox "�Ϸ�"

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

