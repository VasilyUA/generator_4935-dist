Attribute VB_Name = "NewMacros"
Sub GenerateFinalPythonObject()
    Dim folderPath As String, fileName As String, fileNum As String, dayPart As String
    Dim resultText As String, i As Integer
    Dim batDict As Object, bzBatDict As Object
    Set batDict = CreateObject("Scripting.Dictionary")
    Set bzBatDict = CreateObject("Scripting.Dictionary")

    ' Кодування українських назв ключів (щоб не було знаків питань)
    Dim k_bat As String, k_brg As String, k_bz_bat As String, k_bz_brg As String
    k_bat = ChrW(1073) & ChrW(1072) & ChrW(1090) ' "бат"
    k_brg = ChrW(1073) & ChrW(1088) & ChrW(1075) ' "брг"
    k_bz_bat = ChrW(1073) & ChrW(1079) & "_" & k_bat ' "бз_бат"
    k_bz_brg = ChrW(1073) & ChrW(1079) & "_" & k_brg ' "бз_брг"

    ' 1. Вибір папки
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Оберіть папку з файлами"
        If .Show = -1 Then folderPath = .SelectedItems(1) & "\" Else Exit Sub
    End With

    ' 2. Зчитування файлів
    fileName = Dir(folderPath & "*.*")
    Do While fileName <> ""
        If InStr(fileName, ChrW(8470)) > 0 Then ' Пошук знака №
            ' Витягуємо номер
            fileNum = Mid(fileName, InStr(fileName, ChrW(8470)) + 1)
            fileNum = Trim(Split(fileNum, " ")(0))
            
            ' Витягуємо день (шукаємо перші дві цифри дати перед .2026 або подібним)
            ' Шукаємо останню крапку (перед розширенням) і відраховуємо назад
            Dim posLastDot As Integer
            posLastDot = InStrRev(fileName, ".")
            dayPart = Mid(fileName, posLastDot - 10, 2)

            ' Перевірка типу файлу за ключовими словами
            If InStr(1, fileName, "вх", vbTextCompare) > 0 Or InStr(1, fileName, "вих", vbTextCompare) > 0 Then
                batDict(dayPart) = fileNum
            ElseIf InStr(1, fileName, "безпек", vbTextCompare) > 0 Then
                bzBatDict(dayPart) = fileNum
            End If
        End If
        fileName = Dir()
    Loop

    ' 3. Формування тексту
    resultText = "NUMBER_OF_DOCUMENTS = {" & vbCrLf
    
    Dim sPrev As String, sCurr As String
    sPrev = ".{(datetime.strptime(DATA, ""%m.%Y"") - relativedelta(months=1)).strftime(""%m.%Y"")}"
    sCurr = ".{datetime.strptime(DATA, ""%m.%Y"").strftime(""%m.%Y"")}"

    ' Масив для виводу
    Dim days As Variant, currentSuffix As String
    days = Array("29", "30", "31", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "30", "31")

    For i = LBound(days) To UBound(days)
        ' Визначаємо суфікс: для перших трьох елементів (29,30,31) - минулий місяць, далі - поточний
        If i <= 2 Then currentSuffix = sPrev Else currentSuffix = sCurr
        
        Dim d As String: d = days(i)
        Dim v1 As String: v1 = "": If batDict.Exists(d) Then v1 = batDict(d)
        Dim v2 As String: v2 = "": If bzBatDict.Exists(d) Then v2 = bzBatDict(d)

        resultText = resultText & "    f'" & d & currentSuffix & "': " & _
                     "{'" & k_bat & "': '" & v1 & "', '" & k_brg & "': '', '" & _
                     k_bz_bat & "': '" & v2 & "', '" & k_bz_brg & "': ''}," & vbCrLf
    Next i

    resultText = resultText & "}"

    ' 4. Вивід
    Selection.TypeText Text:=resultText
    MsgBox "Готово! Кириличні назви відновлено.", vbInformation
End Sub


