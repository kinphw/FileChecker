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

    i = i + sheet.Comments.Count ' (�����޸�) ����
    i = i + sheet.CommentsThreaded.Count '(Office365�޸�) ����
    
    sheet.UsedRange.ClearComments '�߰� '�����̰� �����̰� �� ����

Next

Dim str As String
str = ActiveWorkbook.Name + "������ ��� �޸�(��Ʈ) ���� �Ϸ�, ������ �޸�� " + CStr(i) + "���Դϴ�."
MsgBox str

End Sub

