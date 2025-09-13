Option Explicit

' --- Модуль v3 (с раскрашиванием строк) ---

Private Const PLAN_SHEET As String = "План"
Private Const PODVESY_SHEET As String = "Подвесы"
Private Const PLAN_TABLE As String = "tblPlan"

Sub ShowTaskForm()
    If Not IsLastRowValid() Then
        Exit Sub ' Don't show the form if validation fails
    End If
    Form1.Show
End Sub

Public Function IsLastRowValid() As Boolean
    IsLastRowValid = True ' Assume valid by default
    Dim lastRowByDate As Long
    Dim trueLastRow As Long
    Dim wsPodvesy As Worksheet
    
    On Error GoTo ValidationFailed

    Set wsPodvesy = ThisWorkbook.Sheets("Подвесы")
    
    ' Find last row based on Date column (D)
    lastRowByDate = wsPodvesy.Cells(wsPodvesy.Rows.Count, "D").End(xlUp).Row
    
    ' Find the very last used row on the sheet
    If Application.WorksheetFunction.CountA(wsPodvesy.Cells) > 0 Then
        trueLastRow = wsPodvesy.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Else
        trueLastRow = 0 ' Sheet is blank
    End If

    ' CHECK 1: Stray data check
    If trueLastRow > lastRowByDate Then
        MsgBox "Обнаружены данные в строке " & trueLastRow & " без указания даты в столбце D. " & _
               "Это приведет к перезаписи данных. Пожалуйста, укажите дату в этой строке или полностью удалите ее.", vbCritical, "Ошибка данных"
        IsLastRowValid = False
        Exit Function
    End If

    ' CHECK 2: Incomplete row check (the existing check on the last row with a date)
    If lastRowByDate < 2 Then Exit Function ' No data rows to check
    
    Dim valG As String, valH As String, valR As String
    valG = Trim(CStr(wsPodvesy.Cells(lastRowByDate, 7).value))
    valH = Trim(CStr(wsPodvesy.Cells(lastRowByDate, 8).value))
    valR = Trim(CStr(wsPodvesy.Cells(lastRowByDate, 18).value))

    If valG = "" Or valH = "" Or valR = "" Then
        MsgBox "В строке " & lastRowByDate & " на листе 'Подвесы' не заполнены все необходимые ячейки (G, H или R). Пожалуйста, заполните их корректно или полностью очистите строку.", vbCritical, "Ошибка данных"
        IsLastRowValid = False
    End If
    
    Exit Function

ValidationFailed:
    Dim errorMsg As String
    errorMsg = "Произошла ошибка при проверке листа 'Подвесы'."
    Dim errorRow As Long
    On Error Resume Next
    errorRow = wsPodvesy.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    On Error GoTo 0
    If errorRow > 0 Then
         errorMsg = "Произошла ошибка при чтении строки " & errorRow & " на листе 'Подвесы'."
    End If

    MsgBox errorMsg & vbCrLf & "Описание: " & Err.Description, vbCritical, "Ошибка проверки"
    IsLastRowValid = False
End Function

Public Sub ExecuteTask(ByVal sheetRowIndex As Long)
    ' Объявление всех переменных в начале процедуры
    Dim wsPlan As Worksheet, wsPodvesy As Worksheet
    Dim tblPlan As ListObject
    Dim targetRow As ListRow, newRow As ListRow
    Dim currentPlanned As Variant, rowData As Variant
    Dim taskIsNowComplete As Boolean
    Dim tempValue As String
    Dim firstGreenIndex As Long
    Dim i As Long
    Const quantity As Long = 1

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set wsPlan = ThisWorkbook.Sheets(PLAN_SHEET)
    
    ' --- Проверка и создание листа "Подвесы" ---
    On Error Resume Next
    Set wsPodvesy = ThisWorkbook.Sheets(PODVESY_SHEET)
    On Error GoTo 0

    If wsPodvesy Is Nothing Then
        Set wsPodvesy = ThisWorkbook.Sheets.Add(After:=wsPlan)
        wsPodvesy.Name = PODVESY_SHEET
        With wsPodvesy
            .Cells(1, 4).value = "Дата"
            .Cells(1, 5).value = "№ по порядку"
            .Cells(1, 6).value = "Время"
            .Cells(1, 7).value = "Смена"
            .Cells(1, 8).value = "Вид материалов"
            .Cells(1, 9).value = "Годность подвеса"
            .Cells(1, 10).value = "Ответственный менеджер"
            .Cells(1, 11).value = "НОМЕР КПЗ"
            .Cells(1, 12).value = "Заказчик"
            .Cells(1, 13).value = "Профиль"
            .Cells(1, 14).value = "Тип завески"
            .Cells(1, 15).value = "Вид обработки"
            .Cells(1, 16).value = "Толщина мкм"
            .Cells(1, 17).value = "Цвет"
            .Cells(1, 18).value = "Подвесы, шт"
            .Cells(1, 19).value = "УСЛОВНЫЕ ПОДВЕСЫ, шт"
            .Cells(1, 20).value = "Ламели, шт"
            .Range("D1:T1").Font.Bold = True
        End With
    End If

    Set tblPlan = wsPlan.ListObjects(PLAN_TABLE)

    For i = 1 To tblPlan.ListRows.Count
        If tblPlan.ListRows(i).Range.Row = sheetRowIndex Then
            Set targetRow = tblPlan.ListRows(i)
            Exit For
        End If
    Next i
    If targetRow Is Nothing Then GoTo ErrorHandler

    ' --- 1. Обновление данных в таблице "План" ---
    taskIsNowComplete = False
    currentPlanned = targetRow.Range.Cells(1, 1).value

    If CStr(currentPlanned) <> "*" Then
        
        Dim cellValue As String: cellValue = CStr(targetRow.Range.Cells(1, 3).value)
        Dim currentCompleted As Long
        Dim textPart As String
        Dim parenPos As Integer
        
        parenPos = InStr(cellValue, "(")
        If parenPos > 0 Then
            currentCompleted = Val(Left(cellValue, parenPos - 1))
            textPart = Mid(cellValue, parenPos)
        Else
            currentCompleted = Val(cellValue)
            textPart = ""
        End If

        If currentCompleted >= CLng(currentPlanned) Then GoTo SkipUpdate

        Dim newCompleted As Long
        newCompleted = currentCompleted + quantity

        If newCompleted >= CLng(currentPlanned) Then
            ' Задача ЗАВЕРШЕНА
            targetRow.Range.Cells(1, 3).value = newCompleted & " (" & Format(Now, "hh:mm dd.mm.yy") & ")"
            taskIsNowComplete = True
        Else
            ' Задача В РАБОТЕ
            If textPart = "" Then
                ' Первый запуск, ПЕРЕМЕЩАЕМ ВВЕРХ (под клюшки)
                tempValue = newCompleted & " (" & Format(Now, "hh:mm dd.mm.yy") & ")"
                targetRow.Range.Cells(1, 3).value = tempValue
                rowData = targetRow.Range.value
                
                targetRow.Delete

                ' --- Новый поиск индекса для вставки (под "клюшками" и задачами в работе) ---
                Dim insertionIndex As Long
                insertionIndex = 1 ' По умолчанию в самый верх
                
                For i = 1 To tblPlan.ListRows.Count
                    Dim isKlyushka As Boolean
                    isKlyushka = (tblPlan.ListRows(i).Range.Cells(1, 4).value = "Клюшки (растрав)")
                    
                    Dim isInProgress As Boolean
                    isInProgress = (tblPlan.ListRows(i).Range.Interior.color = RGB(173, 216, 230))
                    
                    If isKlyushka Or isInProgress Then
                        insertionIndex = i + 1
                    End If
                Next i
                
                Set newRow = tblPlan.ListRows.Add(insertionIndex)
                newRow.Range.value = rowData
                newRow.Range.Interior.color = RGB(173, 216, 230)
                
                Set targetRow = newRow
            Else
                ' Уже в работе
                targetRow.Range.Cells(1, 3).value = newCompleted & " " & textPart
                targetRow.Range.Interior.color = RGB(173, 216, 230)
            End If
        End If
    End If

SkipUpdate:
    ' --- 2. Добавление записи на лист "Подвесы" ---
    Dim lastPodvesyRow As Long, newPodvesyRow As Long
    Dim currentShift As String
    Dim sequenceNumber As Long

    lastPodvesyRow = wsPodvesy.Cells(wsPodvesy.Rows.Count, "D").End(xlUp).Row
    newPodvesyRow = lastPodvesyRow + 1

    If Hour(Now) >= 8 And Hour(Now) < 20 Then
        currentShift = "1-я"
    Else
        currentShift = "2-я"
    End If

    If lastPodvesyRow < 2 Then
        sequenceNumber = 1
    Else
        If currentShift = CStr(wsPodvesy.Cells(lastPodvesyRow, 7).value) Then
            sequenceNumber = wsPodvesy.Cells(lastPodvesyRow, 5).value + 1
        Else
            sequenceNumber = 1
        End If
    End If

    With wsPodvesy
        .Cells(newPodvesyRow, 4).value = Date
        .Cells(newPodvesyRow, 5).value = sequenceNumber
        If Form1.chkRecordTime.value = True Then
            .Cells(newPodvesyRow, 6).value = Format(Time, "hh:mm")
        End If
        .Cells(newPodvesyRow, 7).value = currentShift
        .Cells(newPodvesyRow, 8).value = targetRow.Range.Cells(1, 4).value
        .Cells(newPodvesyRow, 9).value = targetRow.Range.Cells(1, 5).value
        .Cells(newPodvesyRow, 10).value = targetRow.Range.Cells(1, 6).value
        .Cells(newPodvesyRow, 11).value = targetRow.Range.Cells(1, 7).value
        .Cells(newPodvesyRow, 12).value = targetRow.Range.Cells(1, 8).value
        .Cells(newPodvesyRow, 13).value = targetRow.Range.Cells(1, 9).value
        .Cells(newPodvesyRow, 14).value = targetRow.Range.Cells(1, 10).value
        .Cells(newPodvesyRow, 15).value = targetRow.Range.Cells(1, 11).value
        .Cells(newPodvesyRow, 16).value = targetRow.Range.Cells(1, 12).value
        .Cells(newPodvesyRow, 17).value = targetRow.Range.Cells(1, 13).value
        .Cells(newPodvesyRow, 19).value = targetRow.Range.Cells(1, 15).value
        .Cells(newPodvesyRow, 20).value = targetRow.Range.Cells(1, 16).value
        .Cells(newPodvesyRow, 18).value = quantity
    End With

    ' --- 3. Перемещение выполненной задачи в конец таблицы "План" ---
    If taskIsNowComplete Then
        rowData = targetRow.Range.value
        targetRow.Delete

        firstGreenIndex = 0
        For i = 1 To tblPlan.ListRows.Count
            If tblPlan.ListRows(i).Range.Interior.color = RGB(204, 255, 204) Then
                firstGreenIndex = i
                Exit For
            End If
        Next i

        If firstGreenIndex = 0 Then
            tblPlan.ListRows.Add
            Set newRow = tblPlan.ListRows.Add
        Else
            Set newRow = tblPlan.ListRows.Add(firstGreenIndex)
        End If

        newRow.Range.value = rowData
        With newRow.Range
            .Interior.color = RGB(204, 255, 204)
        End With
    End If

ErrorHandler:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Возникла ошибка в Module2: " & Err.Description, vbCritical

End Sub