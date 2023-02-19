Attribute VB_Name = "KillName"
Sub 삭제()

Application.ScreenUpdating = False
Application.EnableEvents = False
Application.Calculation = xlCalculationManual


On Error Resume Next

'For Each c In ThisWorkbook.Names
For Each c In ActiveWorkbook.Names 'ActiveWorkbook으로 해야 작동함 (모듈설정된 추가기능.xlam이 아니니까)
    c.Delete
Next c


MsgBox "완료"

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic

End Sub

