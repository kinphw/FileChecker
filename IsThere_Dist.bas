Attribute VB_Name = "IsThere_Dist"
Function IsThere(a As Range, b As Range)

'�μ� a: VLookup ���
'�μ� b: VLookup ����

'a = ActiveSheet.Range("A1")
'b = ActiveSheet.Range("B1:B2")

'�߰�����
Dim v1 As Variant
Dim v2 As Variant

'1. Vlookup �ǽ� ����� = b

v1 = Application.VLookup(a, b, 1, 0) '�������� �˼� ������ �̰ŷ�
v2 = Not WorksheetFunction.IsError(v1)

IsThere = v2
'
'MsgBox v2

End Function
