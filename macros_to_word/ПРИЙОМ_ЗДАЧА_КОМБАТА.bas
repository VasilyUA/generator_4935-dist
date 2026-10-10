Attribute VB_Name = "Module1"
Sub Collect_B_K_From_All_Sheets()

    Dim ws As Worksheet
    Dim tgt As Worksheet
    Dim lastRow As Long, tRow As Long
    Dim i As Long
    Dim valB As String, valK As String
    
    'Створюємо аркуш Зведення
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Зведення").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set tgt = Worksheets.Add
    tgt.Name = "Зведення"
    
    'Заголовки
    tgt.Range("A1").Value = "Найменування матеріальних засобів"
    tgt.Range("B1").Value = "Одиниця виміру"
    tgt.Range("C1").Value = "Рахується"
    tgt.Range("D1").Value = "В наявності"
    tRow = 2
    
    'Проходимо по всіх аркушах
    For Each ws In Worksheets
        If ws.Name <> "Зведення" Then
            
            lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
            
            For i = 2 To lastRow

                '? Ігноруємо 28-й рядок на всіх аркушах
                If i > 28 Then
                
                    valA = Trim(ws.Cells(i, "A").Value)
                    
                    
                    valB = Trim(ws.Cells(i, "B").Value) 'НАЙМЕНУВАННЯ
                    valI = Trim(ws.Cells(i, "I").Value) 'ОДИНИЦЯ ВИМІРУ
                    valN = Trim(ws.Cells(i, "N").Value) 'РАХУЄТЬСЯ
                    valK = Trim(ws.Cells(i, "K").Value) 'КІЛЬКІСТЬ - В НАЯВНОСТІ
                    
                    'Пропускаємо заголовки
                    If valB <> "" And valK <> "" And valI <> "" And valN <> "" Then
                        If valB <> "Найменування, cтисла характеристика та призначення об’єкта" _
                           And valK <> "Фактична наявність, кіль- кість та (підпис)" _
                           And Not IsNumeric(valB) Then
                            If valA <> "Разом:" Then
                                
                                    tgt.Cells(tRow, "A").Value = valB 'НАЙМЕНУВАННЯ 1
                                    tgt.Cells(tRow, "B").Value = valI 'ОДИНИЦЯ ВИМІРУ 2
                                    tgt.Cells(tRow, "C").Value = valN 'РАХУЄТЬСЯ 3
                                    tgt.Cells(tRow, "D").Value = valK 'КІЛЬКІСТЬ - В НАЯВНОСТІ 4
                                    tRow = tRow + 1
                          End If
                        End If
                    End If
                End If
            Next i
            
        End If
    Next ws
    
    MsgBox "Все зібрано, козаче. 28-й рядок і службові шапки пропущені.", vbInformation

End Sub


