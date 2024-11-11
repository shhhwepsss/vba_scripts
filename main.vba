Sub Подсветить_номера_в_белый_или_желтый()
    Dim tbl As Table
    Dim cell As cell
    Dim targetColumn As Integer
    Dim numbersToHighlight As Variant
    Dim i As Integer
    Dim userInput As String
    Dim cellText As Variant
    Dim j As Integer
    
    ' Укажите номер колонки, в которой нужно искать числа
    targetColumn = 1  ' например, 1 — это первая колонка таблицы

    ' Запрашиваем у пользователя числа для выделения в желтый
    userInputYellow = InputBox("Введите числа для выделения в желтый, разделяя их запятыми (например: 5,16,20):", "Числа для выделения")
    
    
    ' Запрашиваем у пользователя числа для выделения в белый
    userInputWhite = InputBox("Введите числа для выделения в белый, разделяя их запятыми (например: 5,16,20):", "Числа для выделения")
    
    ' Преобразуем строку ввода в массив чисел для желтых
    numbersToHighlightYellow = Split(userInputYellow, " ")
    ' Преобразуем строку ввода в массив чисел белых
    numbersToHighlightWhite = Split(userInputWhite, " ")
    
    ' Проверяем, что в документе есть хотя бы одна таблица
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "В документе нет таблиц."
        Exit Sub
    End If
    
    ' Работаем с первой таблицей в документе
    Set tbl = ActiveDocument.Tables(1)
    
    ' Проходим по каждой строке таблицы в указанной колонке
    For i = 1 To tbl.Rows.Count
        Set cell = tbl.cell(i, targetColumn)
        
        cellText = Val(cell.Range.Text)
        ' Проверяем каждое введенное число на точное совпадение
        For j = LBound(numbersToHighlightYellow) To UBound(numbersToHighlightYellow)
            valueToYellow = Val(Trim(numbersToHighlightYellow(j)))
           shouldPaitnYellow = cellText = valueToYellow
            If shouldPaitnYellow Then
                ' Если найдено точное совпадение, выделяем текст желтым
                cell.Range.HighlightColorIndex = wdYellow
            ElseIf shouldPaitnYellow Then
                cell.Range.HighlightColorIndex = wdWhite
            End If
        Next j
        For K = LBound(numbersToHighlightWhite) To UBound(numbersToHighlightWhite)
            valueToWhite = Val(Trim(numbersToHighlightWhite(K)))
            shouldPaitnWhite = cellText = valueToWhite
            If shouldPaitnWhite Then
                cell.Range.HighlightColorIndex = wdWhite
            End If
        Next K
    Next i
    
    ActiveDocument.Save

    MsgBox "Выделение чисел завершено!"
End Sub

Sub Проверить_Правильно_Ли_Идут_Номера()
    Dim tbl As Table
    Dim cell As cell
    Dim i As Integer
    Dim currentNumber As Integer
    Dim nextNumber As Integer
    Dim report As String
    Dim userMessage As String
    Dim targetColumn As Integer
    
    ' Установите номер столбца для проверки (например, 1 для первого столбца)
    targetColumn = 1
    
    ' Проверяем, есть ли таблицы в документе
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "В документе нет таблиц."
        Exit Sub
    End If
    
    ' Работаем с первой таблицей
    Set tbl = ActiveDocument.Tables(1)
    
    ' Инициализируем строку для отчета
    report = ""
    countOfRowsWithoutLast = tbl.Rows.Count - 1
    
    ' Проходим по строкам таблицы
    For i = 1 To countOfRowsWithoutLast ' Теперь проходим по всем строкам, не включая последнюю
        Set cell = tbl.cell(i, targetColumn)
        
        ' Получаем число в текущей ячейке
        currentNumber = Val(cell.Range.Text)
        
        ' Проверяем, если это не последняя строка, то сравниваем с числом в следующей строке
        If i < countOfRowsWithoutLast Then
            Set cell = tbl.cell(i + 1, targetColumn)
            nextNumber = Val(cell.Range.Text)
            
            ' Проверяем, что следующее число должно быть на единицу больше
            If nextNumber <> currentNumber + 1 Then
                ' Если порядок нарушен, добавляем информацию в отчет
                report = report & "Идет неправильный номер после числа " & currentNumber & ". Ожидалось " & currentNumber + 1 & ", идет " & nextNumber & vbCrLf
            End If
        End If
    Next i
    
    ' Если есть нарушения порядка, выводим их в диалоговом окне
    If report <> "" Then
        userMessage = "Нарушения порядка чисел:" & vbCrLf & report
        MsgBox userMessage
    Else
        MsgBox "Порядок чисел правильный!"
    End If
End Sub
Sub Скорректировать_номера()
    Dim tbl As Table
    Dim cell As cell
    Dim i As Integer
    Dim currentNumber As Integer
    Dim expectedNumber As Integer
    Dim targetColumn As Integer
    
    ' Устанавливаем номер столбца для проверки (например, 2 для второго столбца)
    targetColumn = 1
    
    ' Проверяем, есть ли таблицы в документе
    If ActiveDocument.Tables.Count = 0 Then
        MsgBox "В документе нет таблиц."
        Exit Sub
    End If
    
    ' Работаем с первой таблицей
    Set tbl = ActiveDocument.Tables(1)
    
    ' Устанавливаем ожидаемое первое число (обычно это 1 для второго столбца)
    expectedNumber = 1
    
    ' Проходим по строкам таблицы, начиная с 2-й строки (первая строка с текстом)
    For i = 2 To tbl.Rows.Count ' ???????? ? 2-? ??????
        Set cell = tbl.cell(i, targetColumn)
        
        ' Получаем число в текущей ячейке
        currentNumber = Val(cell.Range.Text)
        
        ' Проверяем, соответствует ли текущее число ожидаемому
        If currentNumber <> expectedNumber Then
            ' Если не соответствует, заменяем его на правильное
            cell.Range.Text = expectedNumber
        End If
        
        ' Обновляем ожидаемое число для следующей строки
        expectedNumber = expectedNumber + 1
    Next i
    
    MsgBox "Числа в таблице были приведены в правильный порядок, начиная с 2-й строки!"
End Sub


