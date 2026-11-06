Attribute VB_Name = "Module1"
Option Explicit

' Замінює координати виду "37U CP 27585 29608" на "*** ** ***** *****"
' у всьому документі (текст, таблиці, колонтитули, написи, виноски),
' не чіпаючи форматування.

Private Const COORD_MASK As String = "*** ** ***** *****"

Sub MaskCoordinates()
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "\b\d{2}[A-Z][ \xA0]*[A-Z]{2}[ \xA0]*\d{5}[ \xA0]+\d{5}\b"
    End With

    Dim total As Long
    Dim story As Range
    Dim rng As Range
    For Each story In ActiveDocument.StoryRanges
        Set rng = story
        Do While Not rng Is Nothing
            total = total + MaskInRange(rng, regEx)
            Set rng = rng.NextStoryRange
        Loop
    Next story

    MsgBox "Готово: замасковано координат - " & total, vbInformation
End Sub

Private Function MaskInRange(ByVal rng As Range, ByVal regEx As Object) As Long
    Dim matches As Object
    Set matches = regEx.Execute(rng.Text)
    If matches.Count = 0 Then Exit Function

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")

    Dim m As Object
    For Each m In matches
        If Not seen.Exists(m.Value) Then
            seen.Add m.Value, True
            ReplaceAllInRange rng, m.Value
        End If
    Next m

    MaskInRange = matches.Count
End Function

Private Sub ReplaceAllInRange(ByVal rng As Range, ByVal findText As String)
    Dim r As Range
    Set r = rng.Duplicate
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = COORD_MASK
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub
