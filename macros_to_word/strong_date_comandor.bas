Attribute VB_Name = "Module1"
Sub BoldAfterKomandyru()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim rngPara As Range
    Dim rngFind As Range
    Dim keyWords As Variant
    Dim kw As Variant
    Dim text As String
    Dim posKom As Integer
    Dim posWord As Integer
    Dim i As Integer

    keyWords = Array("взаб", "1мінбатр", "2мінбатр", "1рбо", "2рбо", "3рбо", "рубак", "рвп", "ісв", "вз", "мп", "рв")

    folderPath = InputBox("Введіть повний шлях до папки з файлами Word:")
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    fileName = Dir(folderPath & "*.doc*")

    Application.ScreenUpdating = False

    While fileName <> ""
        Set doc = Documents.Open(folderPath & fileName, ReadOnly:=False)
        For i = 11 To doc.Paragraphs.Count
            Set rngPara = doc.Paragraphs(i).Range
            text = CleanText(CStr(rngPara.text))
            If Len(text) > 1 Then
                posKom = InStr(1, text, "Командиру ", vbTextCompare)
                If posKom > 0 Then
                    For Each kw In keyWords
                        posWord = InStr(posKom + 9, text, kw, vbTextCompare)
                        If posWord = posKom + 10 Then
                            Set rngFind = doc.Range(rngPara.Start + posWord - 1, rngPara.Start + posWord - 1 + Len(kw))
                            rngFind.Font.Bold = True
                        End If
                    Next
                End If
            End If
        Next i
        doc.Save
        doc.Close
        fileName = Dir
    Wend

    Application.ScreenUpdating = True
    MsgBox "Обробку завершено!"
End Sub

Function CleanText(strIn As String) As String
    Dim s As String
    s = Replace(strIn, Chr(13), "")
    s = Replace(s, Chr(11), "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(160), " ")
    CleanText = s
End Function
