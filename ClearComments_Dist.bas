Attribute VB_Name = "ClearComments_Dist"
Option Explicit

Sub ClearComments(control As IRibbonControl)

Dim i As Integer
i = 0

Dim sheet As Worksheet
Dim xComment As Variant

For Each sheet In Application.ActiveWorkbook.Sheets

'    For Each xComment In xWs.Comments
'        xComment.Delete
'        MsgBox i
'        i = i + 1
'    Next

    i = i + sheet.Comments.Count ' (기존메모) 갯수
    i = i + sheet.CommentsThreaded.Count '(Office365메모) 갯수
    
    sheet.UsedRange.ClearComments '추가 '기존이고 나발이고 다 삭제

Next

Dim str As String
str = ActiveWorkbook.Name + "파일의 모든 메모(노트) 삭제 완료, 삭제한 메모는 " + CStr(i) + "개입니다."
MsgBox str

End Sub

