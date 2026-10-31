Attribute VB_Name = "Module1"
Sub Create_Battle_Table_Human()
    Dim ws As Worksheet
    Dim headers As Variant, subunits As Variant
    Dim i As Long, hdrIndex As Long
    Dim shName As String
    Dim colLetter As String
    Dim currentCol As Long
    Dim rng As Range
    
    Dim srcWS As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim cell As Range
    Dim arr() As String
    Dim idx As Long
    
    Dim dictAttached As Object
    Dim wsAttached As Worksheet
    Dim lastRowAttached As Long
    Dim cellA As Range
    Dim key As Variant
    
    Dim prIndex As Long
    Dim prRow As Long
    Dim endDataRow As Long
    
    Application.ScreenUpdating = False
    shName = "БЧС 2"
    
    ' === видаляємо старий аркуш ===
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets(shName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' === створюємо новий ===
    Set ws = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    ws.Name = shName
    
    ' === головний заголовок ===
    ws.Range("A1").Value = "Бойовий та чисельний склад 1ББО 40 ОБрБО станом на " & _
        Format(Date, "dd.mm.yyyy") & " о " & Format(Time, "HH:MM")
    ws.Range("A1").Font.Size = 10
    ws.Range("A1").Font.Bold = True
    
    ' === масив заголовків ===
    headers = Array("Підрозділи", "За штатом", "За списком", "ПЕРЕВІРКА", "В НАЯВНОСТІ", "% укомпл.", "Район ведення БД", _
                    "Виконують завдання РОП,ВОП,ПВ,СП", "Резерв", "Пілоти", "ВП", "Екіпажі", "ОКП/ТакКП(КСП)", _
                    "Евакуація", "Провідники", "Облаштування фортифікаційних споруд", "Документація / розслідування", _
                    "Отримання та підвоз до позицій", "Забезпечення КСП, ПУВБ", "Контузія", "Потребують ВЛК", _
                    "Обмежено придатні", "Поза межами БД", "ППД", "ТИЛ", "ПУБ", "Шпиталь", "Відрядження", "Відпустка", _
                    "СЗЧ", "Полон", "Безвісті зниклі", "200", "Смерть")
    
    ' === перший заголовок — одна колонка (A2:A3) ===
    ws.Range("A2:A3").Merge
    ws.Range("A2").Value = headers(0)
    ws.Range("A2").HorizontalAlignment = xlCenter
    ws.Range("A2").VerticalAlignment = xlCenter
    ws.Range("A:A").VerticalAlignment = xlCenter
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 10
    
    ' === решта заголовків — по 2 колонки кожен ===
    currentCol = 2
    Dim nextLetter As String
    For hdrIndex = 1 To UBound(headers)
        colLetter = Split(ws.Cells(1, currentCol).Address(True, False), "$")(0)
        nextLetter = Split(ws.Cells(1, currentCol + 1).Address(True, False), "$")(0)
    
        ws.Range(colLetter & "2:" & nextLetter & "2").Merge
        ws.Range(colLetter & "2").Value = headers(hdrIndex)
        ws.Range(colLetter & "2").HorizontalAlignment = xlCenter
        ws.Range(colLetter & "2").VerticalAlignment = xlCenter
        ws.Range(colLetter & "2").Font.Bold = True
        ws.Range(colLetter & "2").Font.Size = 9
    
        ws.Range(colLetter & "3").Value = "Всього"
        ws.Range(nextLetter & "3").Value = "у т.ч. офіцери"
        ws.Range(colLetter & "3:" & nextLetter & "3").HorizontalAlignment = xlCenter
        ws.Range(colLetter & "3:" & nextLetter & "3").VerticalAlignment = xlCenter
        ws.Range(colLetter & "3:" & nextLetter & "3").Font.Size = 9
        ws.Range(colLetter & "3:" & nextLetter & "3").Font.Bold = True
    
        currentCol = currentCol + 2
    Next hdrIndex
    
    ' === обертаємо текст у заголовках ===
    ws.Range("B2:" & nextLetter & "3").Orientation = 90
    ws.Range("B2:" & nextLetter & "3").WrapText = True
    ws.Range("B2:" & nextLetter & "3").HorizontalAlignment = xlCenter
    ws.Range("B2:" & nextLetter & "3").VerticalAlignment = xlCenter
    
    
    ' === список підрозділів з аркуша "40 ОБрБО" ===
    Set srcWS = ThisWorkbook.Sheets("40 ОБрБО")
    Set dict = CreateObject("Scripting.Dictionary")
    
    lastRow = srcWS.Cells(srcWS.Rows.Count, "H").End(xlUp).Row
    For Each cell In srcWS.Range("H2:H" & lastRow)
        If Trim(cell.Value) <> "" Then
            If Not dict.exists(Trim(cell.Value)) Then
                dict.Add Trim(cell.Value), 1
            End If
        End If
    Next cell
    
    ReDim arr(0 To dict.Count)
    arr(0) = "Всього за 1 батальйон"
    idx = 1
    For Each key In dict.keys
        arr(idx) = key
        idx = idx + 1
    Next key
    
    ' === додаємо "Придані" та унікальні значення з аркуша "ПРИДАНІ" ===
    Set dictAttached = CreateObject("Scripting.Dictionary")
    Set wsAttached = ThisWorkbook.Sheets("ПРИДАНІ")
    lastRowAttached = wsAttached.Cells(wsAttached.Rows.Count, "B").End(xlUp).Row
    
    ' додаємо "Придані" як окремий пункт
    ReDim Preserve arr(0 To UBound(arr) + 1)
    arr(UBound(arr)) = "Придані"
    
    For Each cellA In wsAttached.Range("B2:B" & lastRowAttached)
        If Trim(cellA.Value) <> "" Then
            If Not dictAttached.exists(Trim(cellA.Value)) Then
                dictAttached.Add Trim(cellA.Value), 1
                ReDim Preserve arr(0 To UBound(arr) + 1)
                arr(UBound(arr)) = Trim(cellA.Value)
            End If
        End If
    Next cellA
    
    subunits = arr
    
    ' === знайдемо індекс "Придані" у масиві subunits (якщо є) ===
    prIndex = -1
    For i = 0 To UBound(subunits)
        If Trim(subunits(i)) = "Придані" Then
            prIndex = i
            Exit For
        End If
    Next i
    
    If prIndex <> -1 Then
        ' рядок у листі, де стоїть "Придані"
        prRow = 4 + prIndex    ' тому що subunits(0) записується в A4
        ' дані для сумування — від рядка 5 до рядка перед "Придані"
        endDataRow = prRow
        If endDataRow < 5 Then endDataRow = 5 ' мінімальний захист
    Else
        ' fallback (як було раніше)
        endDataRow = 22
    End If
    
    ' === підфарбування блоків  ===
    ws.Range("A2:A4").Interior.Color = RGB(255, 0, 0)                 ' Червоний
    ws.Range("B2:K29").Interior.Color = RGB(146, 208, 80)             ' зелений блок
    ws.Range("L2:M29").Interior.Color = RGB(142, 169, 238)            ' синій
    ws.Range("N2:AQ29").Interior.Color = RGB(198, 239, 206)           ' зелено-блакитний
    ws.Range("AR2:" & nextLetter & "29").Interior.Color = RGB(198, 89, 17)  ' Коричневий
    ws.Range("AT2:" & nextLetter & "29").Interior.Color = RGB(237, 125, 49)  ' Помаранчеви
    ws.Range("B4:BO4").Interior.Color = RGB(255, 255, 0)              ' жовтий
    
    For i = 0 To UBound(subunits)
        ws.Range("A" & i + 4).Value = subunits(i)
        ws.Range("A" & i + 4).Font.Size = 9
        ws.Range("A" & i + 4).HorizontalAlignment = xlLeft
        
        ' фарбуємо червоним, якщо це рядок "Придані"
        If Trim(subunits(i)) = "Придані" Then
            ws.Range("A" & i + 4).Interior.Color = RGB(255, 0, 0)
            ws.Range("B" & i + 4 & ":" & "BO" & i + 4).Interior.Color = RGB(255, 255, 0)
        End If
    Next i
    
    ' === рамки для всієї таблиці ===
    ws.Range("A2:" & nextLetter & UBound(subunits) + 4).Borders.LineStyle = xlContinuous
    
    ' === Рядок 4: автоматичні формули підсумків ===
    Dim lastSummaryCol As Long
    Dim col As Long
    Dim colL As Long, colM As Long
    
    ' знайдемо останню колонку
    lastSummaryCol = ws.Range(nextLetter & "2").Column
    
    ' проставляємо SUM формули по всіх колонках
    For col = 2 To lastSummaryCol
        Dim letter As String
        letter = Split(ws.Cells(4, col).Address(True, False), "$")(0)
    
        ' % укомпл. обробляємо окремо
        If ws.Cells(2, col).Value = "% укомпл." Then
            colL = col
            colM = col + 1
            
            ' використовуємо FormulaLocal для локальної IFERROR( ; )
            ws.Cells(4, colL).FormulaLocal = "=IFERROR(" & ws.Cells(4, colL - 6).Address(False, False) & "/" & ws.Cells(4, colL - 8).Address(False, False) & ";0)"
            ws.Cells(4, colM).FormulaLocal = "=IFERROR(" & ws.Cells(4, colM - 6).Address(False, False) & "/" & ws.Cells(4, colM - 8).Address(False, False) & ";0)"
            
            ' використовуємо FormulaLocal для ПРИДАНІ локальної IFERROR( ; )
            ws.Cells(endDataRow, colL).FormulaLocal = "=IFERROR(" & ws.Cells(endDataRow, colL - 6).Address(False, False) & "/" & ws.Cells(endDataRow, colL - 8).Address(False, False) & ";0)"
            ws.Cells(endDataRow, colM).FormulaLocal = "=IFERROR(" & ws.Cells(endDataRow, colM - 6).Address(False, False) & "/" & ws.Cells(endDataRow, colM - 8).Address(False, False) & ";0)"
            
            col = col + 1 ' перескочити другу колонку % (бо парна)
        Else
            ' стандартна SUM формула — endDataRow визначений динамічно
            ws.Cells(4, col).Formula = "=SUM(" & letter & "5:" & letter & endDataRow & ")"
            ws.Cells(endDataRow, col).Formula = "=SUM(" & letter & "5:" & letter & endDataRow & ")"
        End If
    Next col
    
    ' === розміри ===
    ws.Columns("A").ColumnWidth = 30
    ws.Columns("B:" & nextLetter).ColumnWidth = 1.7
    ws.Rows(1).RowHeight = 25
    ws.Rows(2).RowHeight = 100
    ws.Rows(3).RowHeight = 70
    
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 9
    
    ws.Range("A1:" & nextLetter & "1").Merge
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").VerticalAlignment = xlCenter
    
    Application.ScreenUpdating = True
    MsgBox "БЧС 2 створено!", vbInformation
End Sub



