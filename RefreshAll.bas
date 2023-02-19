Attribute VB_Name = "RefreshAll"
'RefreshAll _ ActiveWorkbook의 모든 피벗을 새로고침


Sub refresh()

 ActiveWorkbook.RefreshAll
 MsgBox ActiveWorkbook.Name + "의 모든 피벗캐시가 새로고침되었습니다."
 
End Sub

