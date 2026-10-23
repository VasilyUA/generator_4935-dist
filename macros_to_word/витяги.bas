Attribute VB_Name = "Module1"
Sub МаскуватиКоординати()
    Dim rng As Range, regex As Object, matches As Object, match As Object
    Dim i As Long, para As Paragraph, doc As Document
    Set doc = ActiveDocument

    ' 1. Розміри полів
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(2)
        .BottomMargin = CentimetersToPoints(2)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(1)
    End With

    ' 1.1 Встановити шрифт для всього документа
    doc.Content.Font.name = "Times New Roman"

    ' 2. Маскування координат
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True: .IgnoreCase = True
        .Pattern = "36T\s+VS\s+(\d{3})\d{2}\s+(\d{3})\d{2}"
    End With

    Set matches = regex.Execute(doc.Content.Text)
    For i = matches.Count - 1 To 0 Step -1
        Set rng = doc.Range(match.FirstIndex, match.FirstIndex + Len(match.Value))
        rng.Text = "36T VS " & matches(i).SubMatches(0) & "** " & matches(i).SubMatches(1) & "**"
        rng.Font.name = "Times New Roman"
    Next i

    ' 3. Видалення службового тексту до ключових слів
    Dim iPara As Long, iStart As Long, iEnd As Long, textPara As String, foundKey As Boolean
    Dim keyWords As Variant: keyWords = Array("командиру", "1рбо", "2рбо", "3рбо", _
        "1 мінбатр", "2 мінбатр", "рвп", "рубак", "рв", "ісв", "вз", "взаб", "мп")

    For iPara = 1 To doc.Paragraphs.Count
        textPara = LCase(Trim(doc.Paragraphs(iPara).Range.Text))
        If InStr(textPara, "для службового користування") > 0 Or _
           InStr(textPara, "прим.№") > 0 Then
            iStart = iPara: Exit For
        End If
    Next iPara

    If iStart > 0 Then
        For iPara = iStart + 1 To doc.Paragraphs.Count
            textPara = LCase(Trim(doc.Paragraphs(iPara).Range.Text))
            For Each Key In keyWords
                If InStr(textPara, Key) > 0 Then iEnd = iPara: Exit For
            Next Key
            If iEnd > 0 Then Exit For
        Next iPara
        If iEnd = 0 Then iEnd = doc.Paragraphs.Count + 1
        For iPara = iEnd - 1 To iStart Step -1
            doc.Paragraphs(iPara).Range.Delete
        Next iPara
    End If

    ' 4. Заміна заголовку
    With doc.Content.Find
        .ClearFormatting: .replacement.ClearFormatting
        .Text = "БОЙОВЕ РОЗПОРЯДЖЕННЯ КОМАНДИРА 1 ББО №"
        .replacement.Text = "ВИТЯГ З БОЙОВОГО РОЗПОРЯДЖЕННЯ КОМАНДИРА 1 ББО №"
        .Execute Replace:=wdReplaceAll
    End With

    ' 5. Вставка блоку підпису
    Dim foundSignature As Boolean: foundSignature = False
    Dim rngInsert As Range: Set rngInsert = doc.Content

    If InStr(doc.Content.Text, "Згідно з оригіналом:") > 0 Then
        foundSignature = True
    End If

    With rngInsert.Find
        .ClearFormatting
        .Text = "Антон СТЕЦЕНКО"
        .Forward = True
        .Wrap = wdFindStop
        If Not foundSignature Then
            If .Execute Then
                InsertSignatureBlock rngInsert, _
                    "Начальник штабу – заступник командира 1ббо 40 обрбо", _
                    "майор", "Тарас ПАСОСЬ"
                foundSignature = True
            End If
        End If
    End With

    If Not foundSignature Then
          For i = 1 To doc.Paragraphs.Count - 1
              If Trim(doc.Paragraphs(i).Range.Text) Like "*ТВО командира 1 ббо 40 обрбо*" Then
                  If InStr(doc.Paragraphs(i + 1).Range.Text, "Тарас ПАСОСЬ") > 0 Then
                      Set rngInsert = doc.Paragraphs(i + 1).Range
                      Call InsertSignatureBlock(rngInsert, _
                          "ТВО начальника штабу – заступника командира 1ббо 40 обрбо", _
                          "молодший лейтенант", "Василь МИРОНЕНКО")
                      foundSignature = True
                      Exit For
                  End If
              End If
          Next i
      End If

    ' 6. ВИДАЛИТИ ВСЕ ПІСЛЯ БЛОКУ "Згідно з оригіналом:"
    Dim iTail As Long, deleteFrom As Long
    deleteFrom = 0

    For iTail = 1 To doc.Paragraphs.Count - 3
        Dim line1 As String, line2 As String, line3 As String, line4 As String
        line1 = LCase(Trim(doc.Paragraphs(iTail).Range.Text))
        line2 = LCase(Trim(doc.Paragraphs(iTail + 1).Range.Text))
        line3 = LCase(Trim(doc.Paragraphs(iTail + 2).Range.Text))
        line4 = LCase(Trim(doc.Paragraphs(iTail + 3).Range.Text))

        If InStr(line1, "згідно з оригіналом") > 0 And _
           (InStr(line2, "начальник штабу") > 0 Or InStr(line2, "тво начальника штабу") > 0) And _
           (InStr(line3, "молодший лейтенант") > 0 Or InStr(line3, "майор") > 0) And _
           (InStr(line3, "василь мироненко") > 0 Or InStr(line3, "тарас пасось") > 0) And _
           InStr(line4, "б/п") > 0 Then
            deleteFrom = iTail + 4
            Exit For
        End If
    Next iTail

    If deleteFrom > 0 And deleteFrom <= doc.Paragraphs.Count Then
        For iTail = doc.Paragraphs.Count To deleteFrom Step -1
            doc.Paragraphs(iTail).Range.Delete
        Next iTail
    End If

    ' 7. Встановити шрифт Times New Roman для всіх параграфів (перестрахування)
    For Each para In doc.Paragraphs
        para.Range.Font.name = "Times New Roman"
    Next para

    MsgBox "Макрос виконано", vbInformation
End Sub

Private Sub InsertSignatureBlock(rngInsert As Range, roleLine As String, rank As String, name As String)
    rngInsert.Collapse Direction:=wdCollapseEnd
    rngInsert.InsertParagraphAfter
    rngInsert.Move unit:=wdParagraph, Count:=1
    rngInsert.Collapse Direction:=wdCollapseStart

    Dim blockText As String
    blockText = vbCrLf & "Згідно з оригіналом:" & vbCrLf & _
                roleLine & vbCrLf & _
                rank & vbTab & name & vbCrLf & "Б/П"
    rngInsert.Text = blockText

    Dim p As Paragraph
    For Each p In rngInsert.Paragraphs
        With p.Range
            .Font.Size = 14
            .Font.name = "Times New Roman"
        End With
        With p.Format
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
        p.TabStops.ClearAll
        p.TabStops.Add Position:=CentimetersToPoints(16.3), Alignment:=wdAlignTabRight
    Next p
End Sub

