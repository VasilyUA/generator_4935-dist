Attribute VB_Name = "Module2"
Sub ОбробитиВСіФайлиВДиректорії()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document

    ' Введи шлях до папки
    folderPath = InputBox("Введи повний шлях до папки:", "Вибір директорії", "C:\Документи\Накази\")

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    fileName = Dir(folderPath & "*.docx")

    Do While fileName <> ""
        Set doc = Documents.Open(folderPath & fileName)
        
        Call МаскуватиКоординати  ' Запускаємо твій макрос

        doc.Save
        doc.Close False
        fileName = Dir()
    Loop

    MsgBox "Обробка завершена.", vbInformation
End Sub
