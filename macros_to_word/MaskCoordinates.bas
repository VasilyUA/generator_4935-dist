Attribute VB_Name = "Module1"
Sub MaskCoordinates()
    Dim rng As Range
    Set rng = ActiveDocument.Content

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "(\d{2})([A-Z])\s*([A-Z]{2})\s*(\d{5})\s+(\d{5})"
    End With

    Dim matches As Object
    Set matches = regEx.Execute(rng.Text)

    Dim match As Object
    For Each match In matches
        Dim numPart As String
        Dim firstLetter As String
        Dim lastTwo As String

        numPart = match.SubMatches(0)
        firstLetter = match.SubMatches(1)
        lastTwo = match.SubMatches(2)

        Dim replacement As String
        replacement = numPart & firstLetter & " " & lastTwo & " ***** *****"

        rng.Text = Replace(rng.Text, match.Value, replacement)
    Next match

    MsgBox "Готово: координати замасковані"
End Sub
