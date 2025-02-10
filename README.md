Function Similarity(s1 As String, s2 As String) As Double
    Dim i As Long, matches As Long
    Dim len1 As Long, len2 As Long

    ' Проверка на пустые строки
    If Len(s1) = 0 Or Len(s2) = 0 Then
        Similarity = 0
        Exit Function
    End If

    ' Ограничение длины строк (чтобы избежать переполнения)
    If Len(s1) > 1000 Or Len(s2) > 1000 Then
        Similarity = 0
        Exit Function
    End If

    ' Длина строк
    len1 = Len(s1)
    len2 = Len(s2)

    ' Минимальная длина
    Dim minLen As Long
    minLen = Application.Min(len1, len2)

    ' Сравнение символов по порядку
    For i = 1 To minLen
        If Mid(s1, i, 1) = Mid(s2, i, 1) Then
            matches = matches + 1
        End If
    Next i

    ' Процент совпадения
    Similarity = matches / Application.Max(len1, len2)
End Function

Sub CopyValuesWithSimilarity()
    Dim ws As Worksheet
    Dim lastRowD As Long, lastRowG As Long
    Dim i As Long, j As Long
    Dim similarityThreshold As Double
    Dim sim As Double
    Dim found As Boolean

    ' Укажите рабочий лист
    Set ws = ThisWorkbook.Sheets("test123") ' Замените "test123" на имя вашего листа

    ' Найти последнюю заполненную строку в столбцах D и G
    lastRowD = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    lastRowG = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row

    ' Порог схожести (80%)
    similarityThreshold = 0.8

    ' Цикл по значениям столбца D
    For i = 1 To lastRowD
        found = False
        ' Проверка на пустую ячейку
        If Len(ws.Cells(i, "D").Value) > 0 Then
            ' Цикл по значениям столбца G
            For j = 1 To lastRowG
                If Len(ws.Cells(j, "G").Value) > 0 Then
                    sim = Similarity(CStr(ws.Cells(i, "D").Value), CStr(ws.Cells(j, "G").Value))
                    If sim >= similarityThreshold Then
                        ' Если найдено совпадение, копируем из H в B
                        ws.Cells(i, "B").Value = ws.Cells(j, "H").Value
                        found = True
                        Exit For
                    End If
                End If
            Next j
        End If
    Next i

    MsgBox "Копирование завершено!", vbInformation
End Sub
