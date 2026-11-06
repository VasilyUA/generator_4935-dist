Attribute VB_Name = "NewMacros"
Sub GeneratePythonObjectWithStrptime()
    Dim folderPath As String
    Dim fileName As String
    Dim fileNum As String
    Dim fileDate As String
    Dim dataDict As Object
    Set dataDict = CreateObject("Scripting.Dictionary")
    
    Dim i As Integer
    Dim dayStr As String
    Dim resultText As String
    Dim pythonDatePart As String
    
    ' „астина коду Python, €ку ви просили вставити п≥сл€ дн€
    pythonDatePart = ".{datetime.strptime(DATA, ""%m.%Y"").strftime(""%m.%Y"")}"

    ' 1. ¬иб≥р папки
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "ќбер≥ть папку з файлами"
        If .Show = -1 Then
            folderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With

    ' 2. «читуванн€ файл≥в (шукаЇмо номер п≥сл€ є та дату в к≥нц≥)
    fileName = Dir(folderPath & "*«ј¬ƒјЌЌя*.doc*")
    
    Do While fileName <> ""
        If InStr(fileName, "є") > 0 Then
            ' ¬ит€гуЇмо номер бат
            fileNum = Mid(fileName, InStr(fileName, "є") + 1)
            fileNum = Trim(Split(fileNum, " ")(0))
            
            ' ¬ит€гуЇмо т≥льки ƒ≈Ќ№ (перш≥ два знаки дати в назв≥ файлу)
            ' ѕрипускаЇмо формат "... 01.01.2026.docx"
            Dim posDot As Integer
            posDot = InStrRev(fileName, ".") ' крапка перед розширенн€м
            dayStr = Mid(fileName, posDot - 10, 2) ' беремо т≥льки перш≥ 2 символи дати (день)
            
            ' «бер≥гаЇмо в словник ( люч - день, «наченн€ - номер)
            If Not dataDict.Exists(dayStr) Then
                dataDict.Add dayStr, fileNum
            End If
        End If
        fileName = Dir()
    Loop

    ' 3. ‘ормуванн€ тексту об'Їкта
    resultText = "NUMBER_OF_DOCUMENTS_GENERAL = {" & vbCrLf
    
    ' ÷икл суворо в≥д 01 до 30 (або 31)
    For i = 1 To 31
        dayStr = Format(i, "00")
        
        Dim batValue As String
        If dataDict.Exists(dayStr) Then
            batValue = dataDict(dayStr)
        Else
            batValue = "" ' €кщо файлу на цю дату немаЇ
        End If
        
        ' ‘ормуЇмо р€док: f'ƒ≈Ќ№.{datetime...}': {'бат': 'Ќќћ≈–', 'брг': ''},
        resultText = resultText & "    f'" & dayStr & pythonDatePart & "': " & _
                     "{'бат': '" & batValue & "', 'брг': ''}," & vbCrLf
    Next i

    resultText = resultText & "}"

    ' 4. ¬ставка в документ
    Selection.TypeText Text:=resultText
    
    MsgBox "ќб'Їкт сформовано зг≥дно з шаблоном Python!", vbInformation
End Sub


