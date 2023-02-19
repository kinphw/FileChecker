Attribute VB_Name = "Translate"
'1 : ���ڸ���   customButton11
'2 : ������ ��  customButton12
'3 : �Ʒ� ��    customButton13

Public blGo As Boolean

Sub GT_Exe(control As IRibbonControl)

Call Test

If blGo = False Then
    Exit Sub
End If

'ȣ���� ��Ʈ�� ID = ctID
Dim ctID As String
ctID = control.ID

Application.ScreenUpdating = False

'������� => �� ���� ������ ��Ȱ��
Dim str As String
'�������
Dim tmpStr As String
'Ÿ�ټ� (���� ��)
Dim tgtRange As Range

'��ȯ����
For Each rng In Selection

    '1. ������ ����
    str = rng.Value
    tmpStr = GTdo(str)
    
    If (ctID = "customButton11") Then
        str = str + vbCrLf + tmpStr
    Else
        str = tmpStr
    End If
    
    
    '2. ��� ����
    If (ctID = "customButton11") Then
        rng.Value = str
    ElseIf (ctID = "customButton12") Then
        rng.Offset(0, 1).Value = str
    ElseIf (ctID = "customButton13") Then
        rng.Offset(1, 0).Value = str
    End If

Next

Application.ScreenUpdating = True

MsgBox "Done"

End Sub

''''''''''''''''''''''
' ���ϴ� �����Լ���
''''''''''''''''''''''

'�� ���� �׽�Ʈ

Sub Test()

blGo = True

Dim countCells As Long
countCells = Selection.Cells.Count

'���� 100�� �ʰ���
If (countCells > 100) Then
     If MsgBox("������ ���� 100���� �ѽ��ϴ�. ���� �����ü������ �ϸ� VBA�� �ı��ɼ��� �ֽ��ϴ�. �����Ͻðڽ��ϱ�?", vbYesNo) = vbNo Then
    
        blGo = False
     
     End If
End If
End Sub

Function GTinput(str As String, tgt As Range)

'GTinput��
'str�� ���� ����
'tgt�� ���� ����

tgt.Value = str

End Function

Function GTdo(src As String)

'GTdo�� ������ ���ѹ����� ��ȯ��
GTdo = GoogleTranslate(src, "en", "ko")

End Function


Public Function GoogleTranslate(strInput As String, strFromSourceLanguage As String, strToTargetLanguage As String) As String
    Dim strURL As String
    Dim objHTTP As Object
    Dim objHTML As Object
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String

    ' send query to web page
    strURL = "https://translate.google.com/m?hl=" & strFromSourceLanguage & _
        "&sl=" & strFromSourceLanguage & _
        "&tl=" & strToTargetLanguage & _
        "&ie=UTF-8&prev=_m&q=" & strInput

    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP") 'late binding
    objHTTP.Open "GET", strURL, False
    'objHTTP.setRequestHeader "Accept-Encoding", "gzip;q=1.0", "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ""

    'MsgBox objHTTP.getResponseHeader("Content-Encoding")

    Dim httpresponse() As Byte
    httpresponse = objHTTP.responseBody
    'Mod_Inflate64.Inflate httpresponse
    'MsgBox StrConv(httpresponse, vbUnicode)
    'Debug.Print StrConv(httpresponse, vbUnicode)


    ' create an html document
    Set objHTML = CreateObject("htmlfile")
    With objHTML
        .Open
        .Write objHTTP.responseText
        '#YGE 221001 for trialrun --  Z = objHTTP.responseText
        '#YGE 221001 for trialrun --  Y = InStr(Z, "result-container")
        '#YGE 221001 for trialrun --  If Y > 0 Then
        '#YGE 221001 for trialrun --      Debug.Print Z
        '#YGE 221001 for trialrun --  End If
        
        .Close
    End With
    
    '#YGE 221001 for trialrun --  Range("H1") = objHTTP.responsetext
    
    '#YGE 221001 for trialrun --  Set objDivsBody = objHTML.getElementsByTagName("body")
    Set objDivs = objHTML.getElementsByTagName("div")
    '#YGE 221001 for trialrun --  Set objDivs2 = objDivsBody(0).getElementsByTagName("div")
    '#YGE 221001 for trialrun --  Set objSpans = objHTML.getElementsByTagName("span")
    '#YGE 221001 for trialrun --  Set objSpans2 = objDivsBody(0).getElementsByTagName("span")
    
    
    '#YGE 221001 for trialrun --  Set objDivs2 = objHTML.getElementsByClassName("Q4iAWc")
    '#YGE 221001 for trialrun --  Set objDivs2 = objHTML.getElementsByClassName("JLqJ4b ChMk0b")
    
    '#YGE 221001 for trialrun --  For Each objDiv In objDivsBody
        '#YGE 221001 for trialrun --  Z = objDiv.className
        '#YGE 221001 for trialrun --  Debug.Print Z
    '#YGE 221001 for trialrun --  Next objDiv
    
    For Each objDiv In objDivs
        '#YGE 221001 for trialrun --  Z = objDiv.className
        '#YGE 221001 for trialrun --  Debug.Print Z
        
        GoogleTranslate = GoogleTranslateRecursion(objDiv.ChildNodes)
        '#YGE 221001 for trialrun --  Debug.Print GoogleTranslate
        If GoogleTranslate <> "" Then
            Exit For
        End If
        If objDiv.className = "result-container" Then
            strTranslated = objDiv.innerText
            GoogleTranslate = strTranslated
            Exit For
        End If
        
    Next objDiv


    Set objHTML = Nothing
    Set objHTTP = Nothing

End Function


Function GoogleTranslateRecursion(pobjDivs As Object) As String
    Dim objDivs As Object, objDiv As Object
    Dim strTranslated As String
    GoogleTranslateRecursion = ""
    Set objDivs = pobjDivs
    For Each objDiv In objDivs
        If objDiv.nodeName = "DIV" Then
            '#YGE 221001 for trialrun --  Z = objDiv.className
            '#YGE 221001 for trialrun --  Debug.Print Z
            strTranslated = GoogleTranslateRecursion(objDiv.ChildNodes)
            '#YGE 221001 for trialrun --  Debug.Print strTranslated
            If strTranslated <> "" Then
                GoogleTranslateRecursion = strTranslated
                Exit For
            End If
            
            If objDiv.className = "result-container" Then
                strTranslated = objDiv.innerText
                GoogleTranslateRecursion = strTranslated
                Exit For
            End If
        End If
        
    Next objDiv
End Function

