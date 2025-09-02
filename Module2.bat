Option Explicit

' --- Модуль v3 (с раскрашиванием строк) ---

Private Const PLAN_SHEET As String = "План"
Private Const PODVESY_SHEET As String = "Подвесы"
Private Const PLAN_TABLE As String = "tblPlan"

Sub ShowTaskForm()
    Form1.Show
End Sub

Public Sub ExecuteTask(ByVal sheetRowIndex As Long)
    Dim wsPlan As Worksheet, wsPodvesy As Worksheet
    Dim tblPlan As ListObject
    Dim targetRow As ListRow
    Dim currentPlanned As Variant
    Dim currentCompleted As Long
    Dim taskIsNowComplete As Boolean
    Const quantity As Long = 1

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Set wsPlan = ThisWorkbook.Sheets(PLAN_SHEET)
    Set wsPodvesy = ThisWorkbook.Sheets(PODVESY_SHEET)
    Set tblPlan = wsPlan.ListObjects(PLAN_TABLE)

    Dim i As Long
    For i = 1 To tblPlan.ListRows.Count
        If tblPlan.ListRows(i).Range.Row = sheetRowIndex Then
            Set targetRow = tblPlan.ListRows(i)
            Exit For
        End If
    Next i
    If targetRow Is Nothing Then GoTo ErrorHandler

    ' --- 1. Обновление данных в таблице "План" ---
    taskIsNowComplete = False
    With targetRow.Range
        currentPlanned = .Cells(1, 1).Value
        If CStr(currentPlanned) <> "*" Then
            currentCompleted = .Cells(1, 3).Value
            If currentCompleted >= CLng(currentPlanned) Then GoTo SkipUpdate
            .Cells(1, 3).Value = currentCompleted + quantity
            If (.Cells(1, 3).Value) >= CLng(currentPlanned) Then
                taskIsNowComplete = True ' Задача будет перемещена и станет зеленой
            Else
                .Interior.color = RGB(255, 218, 185) ' Светло-оранжевый для задач "В работе"
            End If
        End If
    End With

SkipUpdate:
    ' --- 2. Добавление записи на лист "Подвесы" ---
    Dim lastPodvesyRow As Long
    lastPodvesyRow = wsPodvesy.Cells(wsPodvesy.Rows.Count, "E").End(xlUp).Row + 1
    With wsPodvesy
        .Cells(lastPodvesyRow, 5).Value = targetRow.Range.Cells(1, 4).Value
        .Cells(lastPodvesyRow, 6).Value = targetRow.Range.Cells(1, 5).Value
        .Cells(lastPodvesyRow, 7).Value = targetRow.Range.Cells(1, 6).Value
        .Cells(lastPodvesyRow, 8).Value = targetRow.Range.Cells(1, 7).Value
        .Cells(lastPodvesyRow, 9).Value = targetRow.Range.Cells(1, 8).Value
        .Cells(lastPodvesyRow, 10).Value = targetRow.Range.Cells(1, 9).Value
        .Cells(lastPodvesyRow, 11).Value = targetRow.Range.Cells(1, 10).Value
        .Cells(lastPodvesyRow, 12).Value = targetRow.Range.Cells(1, 11).Value
        .Cells(lastPodvesyRow, 13).Value = targetRow.Range.Cells(1, 12).Value
        .Cells(lastPodvesyRow, 14).Value = targetRow.Range.Cells(1, 13).Value
        .Cells(lastPodvesyRow, 16).Value = targetRow.Range.Cells(1, 15).Value
        .Cells(lastPodvesyRow, 17).Value = targetRow.Range.Cells(1, 16).Value
        .Cells(lastPodvesyRow, 15).Value = quantity
    End With

    ' --- 3. Перемещение выполненной задачи в конец таблицы "План" ---
    If taskIsNowComplete Then
        Dim rowData As Variant
        rowData = targetRow.Range.Value
        targetRow.Delete
        Dim newRow As ListRow
        Set newRow = tblPlan.ListRows.Add
        newRow.Range.Value = rowData
        With newRow.Range
            .Interior.color = RGB(204, 255, 204) ' Светло-зеленый
        End With
    End If

ErrorHandler:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then MsgBox "Возникла ошибка в Module2: " & Err.Description, vbCritical

End Sub
