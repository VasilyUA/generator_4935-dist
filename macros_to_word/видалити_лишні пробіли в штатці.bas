Attribute VB_Name = "Module1"
Option Explicit

Sub CleanExtraSpaces()
    Dim rng As Range
    Dim cell As Range
    Dim txt As String
    Dim re As Object

    ' Створюємо об'єкт RegExp
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = True

    ' Вибраний діапазон
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then
        MsgBox "Виділіть діапазон клітинок для очищення."
        Exit Sub
    End If

    For Each cell In rng
        If Not IsEmpty(cell) And Not cell.HasFormula Then
            txt = CStr(cell.Value)

            ' 1. Замінюємо всі нерозривні пробіли на звичайні
            txt = Replace(txt, Chr(160), " ")

            ' 2. Регулярка: будь-які пробільні символи (space, tab, NBSP, Unicode) у послідовності > один пробіл
            re.Pattern = "\s+"
            txt = re.Replace(txt, " ")

            ' 3. Прибираємо пробіли з початку та кінця
            txt = Trim(txt)

            cell.Value = txt
        End If
    Next cell

    MsgBox "Очищення пробілів завершено!"
End Sub

