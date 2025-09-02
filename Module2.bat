Option Explicit

' --- Модуль v3 (с раскрашиванием строк) ---

Private Const PLAN_SHEET As String = "План"
Private Const PODVESY_SHEET As String = "Подвесы"
Private Const PLAN_TABLE As String = "tblPlan"

Sub ShowTaskForm()
    Form1.Show
End Sub

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
    Static hasShownSheetExistsMsg As Boolean
    On Error Resume Next
    Set wsPodvesy = ThisWorkbook.Sheets(PODVESY_SHEET)
    On Error GoTo 0

    If wsPodvesy Is Nothing Then
        Set wsPodvesy = ThisWorkbook.Sheets.Add(After:=wsPlan)
        wsPodvesy.Name = PODVESY_SHEET
        With wsPodvesy
            .Cells(1, 4).Value = "Дата"
            .Cells(1, 5).Value = "№ по порядку"
            .Cells(1, 6).Value = "Время"
            .Cells(1, 7).Value = "Смена"
            .Cells(1, 8).Value = "Вид материалов"
            .Cells(1, 9).Value = "Годность подвеса"
            .Cells(1, 10).Value = "Ответственный менеджер"
            .Cells(1, 11).Value = "НОМЕР КПЗ"
            .Cells(1, 12).Value = "Заказчик"
            .Cells(1, 13).Value = "Профиль"
            .Cells(1, 14).Value = "Тип завески"
            .Cells(1, 15).Value = "Вид обработки"
            .Cells(1, 16).Value = "Толщина мкм"
            .Cells(1, 17).Value = "Цвет"
            .Cells(1, 18).Value = "Подвесы, шт"
            .Cells(1, 19).Value = "УСЛОВНЫЕ ПОДВЕСЫ, шт"
            .Cells(1, 20).Value = "Ламели, шт"
            .Range("D1:T1").Font.Bold = True
        End With
        hasShownSheetExistsMsg = True
    Else
        If Not hasShownSheetExistsMsg Then
            MsgBox "Лист '" & PODVESY_SHEET & "' уже существует. Работаем с ним.", vbInformation
            hasShownSheetExistsMsg = True
        End If
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
    currentPlanned = targetRow.Range.Cells(1, 1).Value

    If CStr(currentPlanned) <> "*" Then
        
        Dim cellValue As String: cellValue = CStr(targetRow.Range.Cells(1, 3).Value)
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
            targetRow.Range.Cells(1, 3).Value = newCompleted & " (" & Format(Now, "hh:mm dd.mm.yy") & ")"
            taskIsNowComplete = True
        Else
            ' Задача В РАБОТЕ
            If textPart = "" Then
                ' Первый запуск, ПЕРЕМЕЩАЕМ ВВЕРХ (под клюшки)
                tempValue = newCompleted & " (" & Format(Now, "hh:mm dd.mm.yy") & ")"
                targetRow.Range.Cells(1, 3).Value = tempValue
                rowData = targetRow.Range.Value
                
                targetRow.Delete

                ' --- Новый поиск индекса для вставки (под "клюшками" и задачами в работе) ---
                Dim insertionIndex As Long
                insertionIndex = 1 ' По умолчанию в самый верх
                
                For i = 1 To tblPlan.ListRows.Count
                    Dim isKlyushka As Boolean
                    isKlyushka = (tblPlan.ListRows(i).Range.Cells(1, 4).Value = "Клюшки (растрав)")
                    
                    Dim isInProgress As Boolean
                    isInProgress = (tblPlan.ListRows(i).Range.Interior.Color = RGB(173, 216, 230))
                    
                    If isKlyushka Or isInProgress Then
                        insertionIndex = i + 1
                    End If
                Next i
                
                Set newRow = tblPlan.ListRows.Add(insertionIndex)
                newRow.Range.Value = rowData
                newRow.Range.Interior.Color = RGB(173, 216, 230)
                
                Set targetRow = newRow
            Else
                ' Уже в работе
                targetRow.Range.Cells(1, 3).Value = newCompleted & " " & textPart
                targetRow.Range.Interior.Color = RGB(173, 216, 230)
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
        If currentShift = CStr(wsPodvesy.Cells(lastPodvesyRow, 7).Value) Then
            sequenceNumber = wsPodvesy.Cells(lastPodvesyRow, 5).Value + 1
        Else
            sequenceNumber = 1
        End If
    End If

    With wsPodvesy
        .Cells(newPodvesyRow, 4).Value = Date
        .Cells(newPodvesyRow, 5).Value = sequenceNumber
        .Cells(newPodvesyRow, 6).Value = Time
        .Cells(newPodvesyRow, 7).Value = currentShift
        .Cells(newPodvesyRow, 8).Value = targetRow.Range.Cells(1, 4).Value
        .Cells(newPodvesyRow, 9).Value = targetRow.Range.Cells(1, 5).Value
        .Cells(newPodvesyRow, 10).Value = targetRow.Range.Cells(1, 6).Value
        .Cells(newPodvesyRow, 11).Value = targetRow.Range.Cells(1, 7).Value
        .Cells(newPodvesyRow, 12).Value = targetRow.Range.Cells(1, 8).Value
        .Cells(newPodvesyRow, 13).Value = targetRow.Range.Cells(1, 9).Value
        .Cells(newPodvesyRow, 14).Value = targetRow.Range.Cells(1, 10).Value
        .Cells(newPodvesyRow, 15).Value = targetRow.Range.Cells(1, 11).Value
        .Cells(newPodvesyRow, 16).Value = targetRow.Range.Cells(1, 12).Value
        .Cells(newPodvesyRow, 17).Value = targetRow.Range.Cells(1, 13).Value
        .Cells(newPodvesyRow, 19).Value = targetRow.Range.Cells(1, 15).Value
        .Cells(newPodvesyRow, 20).Value = targetRow.Range.Cells(1, 16).Value
        .Cells(newPodvesyRow, 18).Value = quantity
    End With

    ' --- 3. Перемещение выполненной задачи в конец таблицы "План" ---
    If taskIsNowComplete Then
        rowData = targetRow.Range.Value
        targetRow.Delete

        firstGreenIndex = 0
        For i = 1 To tblPlan.ListRows.Count
            If tblPlan.ListRows(i).Range.Interior.Color = RGB(204, 255, 204) Then
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

        newRow.Range.Value = rowData
        With newRow.Range
            .Interior.Color = RGB(204, 255, 204)
        End With
    End If

ErrorHandler:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Возникла ошибка в Module2: " & Err.Description, vbCritical

End Sub
