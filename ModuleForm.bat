Option Explicit

' --- КОД С УЛУЧШЕНИЯМИ: СУММА ПОДВЕСОВ И ЧИСТОЕ ФОРМАТИРОВАНИЕ ---

' === Константы ===
Private Const COL_PLAN As Long = 1
Private Const COL_PRIORITY As Long = 2
Private Const COL_COMPLETED As Long = 3
Private Const COL_CUSTOMER As Long = 8
Private Const COL_PROFILE As Long = 9
Private Const COL_MATERIAL As Long = 4
Private Const COL_KPZ As Long = 7
Private Const COL_COLOR As Long = 13
Private Const COL_LAMELI As Long = 16

Private Sub lblTotal_Click()

End Sub

Private Sub lstTasks_Click()

End Sub


Private Sub UserForm_Click()

End Sub

' --- События ---
Private Sub UserForm_Initialize()
    ' --- Логика загрузки состояния чекбокса из Именованного Диапазона ---
    Dim recordTimeSetting As Name
    
    On Error Resume Next
    Set recordTimeSetting = ThisWorkbook.Names("RecordTimeSetting")
    On Error GoTo 0

    If recordTimeSetting Is Nothing Then
        ' Имя не существует, создаем его со значением по умолчанию TRUE
        ThisWorkbook.Names.Add Name:="RecordTimeSetting", RefersTo:="=TRUE"
        Me.chkRecordTime.value = True
    Else
        ' Имя существует, считываем его значение
        ' RefersTo возвращает строку "=TRUE" или "=FALSE", Evaluate преобразует ее в логическое значение
        Me.chkRecordTime.value = Evaluate(recordTimeSetting.RefersTo)
    End If

    ' --- Основная логика ---
    Call PopulateList
End Sub

Private Sub chkRecordTime_Click()
    ' Сохраняем состояние чекбокса в Именованный Диапазон
    On Error Resume Next ' На случай если имя было удалено
    ThisWorkbook.Names("RecordTimeSetting").RefersTo = "=" & UCase(CStr(Me.chkRecordTime.value))
    On Error GoTo 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' --- Кнопки ---
Private Sub cmdAddTask_Click()
    Call HandleTask(CloseAfter:=False)
End Sub
Private Sub cmdAddAndClose_Click()
    Call HandleTask(CloseAfter:=True)
End Sub

' --- Логика обработки клика ---
Private Sub HandleTask(CloseAfter As Boolean)
    If Me.lstTasks.ListIndex = -1 Then Exit Sub
    
    Dim selectedText As String
    selectedText = Me.lstTasks.value ' Запоминаем текст ДО всех изменений

    If Left(selectedText, 3) = "---" Then Exit Sub
    
    ' --- НОВАЯ УЛУЧШЕННАЯ ЛОГИКА ---
    
    ' 1. Извлекаем "стабильную" часть текста задачи, которая не меняется.
    '    Это будет наш "якорь" для поиска после всех перемещений.
    Dim stableIdentifier As String
    Dim separatorPos As Long
    separatorPos = InStr(selectedText, " - ")
    
    If separatorPos > 0 Then
        stableIdentifier = Mid(selectedText, separatorPos) ' Получаем " - KPZ123 - Заказчик..."
    Else
        stableIdentifier = selectedText ' На случай, если в строке нет дефиса (например, только "*")
    End If

    ' 2. Находим первоначальный номер строки, чтобы передать его в ExecuteTask
    Dim initialRowIndex As Long
    initialRowIndex = FindRowByText(selectedText)
    
    If initialRowIndex > 0 Then
        ' 3. Выполняем задачу. В ЭТОТ МОМЕНТ СТРОКА ПЕРЕМЕЩАЕТСЯ,
        '    и initialRowIndex становится неактуальным!
        Call ExecuteTask(initialRowIndex)
        
        If CloseAfter Then
            Unload Me
            Exit Sub
        End If

        ' Перезагружаем список с учетом всех изменений на листе
        Call PopulateList

        ' 4. ИЩЕМ ЗАДАЧУ В НОВОМ СПИСКЕ ПО СТАБИЛЬНОМУ ИДЕНТИФИКАТОРУ
        Dim i As Long
        Dim isFound As Boolean
        isFound = False
        
        For i = 0 To Me.lstTasks.ListCount - 1
            ' Мы ищем строку, которая ЗАКАНЧИВАЕТСЯ на наш "якорь".
            ' Это сработает, даже если количество в начале изменилось (с 10П на 9П).
            ' Используем CStr для надежности, если в списке окажется не текстовое значение.
            If InStr(CStr(Me.lstTasks.List(i)), stableIdentifier) > 0 Then
                Me.lstTasks.ListIndex = i ' Нашли! Выделяем.
                isFound = True
                Exit For ' Выходим, т.к. задача найдена
            End If
        Next i
        
        ' 5. Если после цикла задача не найдена, значит, она была последней и завершилась.
        If Not isFound Then
            MsgBox "Задача '" & selectedText & "' полностью выполнена и убрана из списка.", vbInformation, "Задача завершена"
            ' Опционально: выделить первую задачу в списке, чтобы фокус не пропадал
            If Me.lstTasks.ListCount > 0 Then Me.lstTasks.ListIndex = 0
        End If
        
        Me.lstTasks.SetFocus
        
    Else
        ' Эта ошибка может возникнуть, если данные на листе изменились между PopulateList и кликом
        MsgBox "Не удалось найти исходную строку на листе: '" & selectedText & "'. Список будет обновлен.", vbExclamation
        Call PopulateList
    End If
End Sub
' --- Основная процедура отображения ---
Public Sub PopulateList()
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim priorityTasks As New Collection
    Dim standardTasks As New Collection
    Dim totalHangersSum As Double ' ИЗМЕНЕНИЕ 2: Переменная для суммы подвесов
    
    On Error GoTo 0

    Set wsPlan = ThisWorkbook.Sheets("План")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row

    Me.lstTasks.Clear
    Me.lstTasks.ColumnCount = 1
    Me.lstTasks.ColumnWidths = "700"
    
    If lastRow < 2 Then
        Me.lstTasks.AddItem "На листе 'План' нет данных."
        Exit Sub
    End If
    
    For i = 2 To lastRow
        Dim planValue As Variant: planValue = wsPlan.Cells(i, COL_PLAN).value
        Dim completedValue As Variant: completedValue = wsPlan.Cells(i, COL_COMPLETED).value
        
        Dim isPriority As Boolean
        Dim priorityCheckValue As Variant
        priorityCheckValue = wsPlan.Cells(i, COL_PRIORITY).value
        
        isPriority = False
        If IsNumeric(priorityCheckValue) Then
            If CDbl(priorityCheckValue) > 0 Then isPriority = True
        End If
        
        Dim remaining As Long
        If CStr(planValue) = "*" Then
            remaining = 999999
        Else
            Dim numPlan As Long, numCompleted As Long
            If IsNumeric(planValue) Then numPlan = CLng(planValue)
            numCompleted = GetCompletedCount(completedValue)
            remaining = numPlan - numCompleted
        End If
        
        If remaining > 0 Then
            ' ИЗМЕНЕНИЕ 2: Суммируем подвесы только для числовых задач
            If CStr(planValue) <> "*" Then
                totalHangersSum = totalHangersSum + remaining
            End If
            
            ' ИЗМЕНЕНИЕ 3: Используем новую функцию для чистого форматирования строки
            Dim displayText As String
            displayText = CreateTaskDisplayString(wsPlan, i)
            
            If isPriority Then
                priorityTasks.Add displayText
            Else
                standardTasks.Add displayText
            End If
        End If
    Next i

    Dim item As Variant
    For Each item In priorityTasks
        Me.lstTasks.AddItem item
    Next item
    
    ' ИЗМЕНЕНИЕ 1: Эта логика уже верна. Разделитель появится, только если есть ОБА типа задач.
    If priorityTasks.Count > 0 And standardTasks.Count > 0 Then
        Me.lstTasks.AddItem "--- СРОЧНО --- СРОЧНО --- СРОЧНО --- СРОЧНО ---"
    End If
    
    For Each item In standardTasks
        Me.lstTasks.AddItem item
    Next item

    Dim totalTasks As Long
    totalTasks = priorityTasks.Count + standardTasks.Count
    
    If totalTasks = 0 Then
        Me.lstTasks.AddItem "Все задачи выполнены."
        Me.Controls("lblTotal").Caption = "Общее кол-во подвесов: 0"
    Else
        ' ИЗМЕНЕНИЕ 2: Обновляем надпись, чтобы показывать сумму подвесов
        Me.Controls("lblTotal").Caption = "Общее кол-во подвесов: " & totalHangersSum
    End If
End Sub

' --- Вспомогательная функция для парсинга количества выполненных ---
Private Function GetCompletedCount(ByVal cellValue As Variant) As Long
    Dim parenPos As Integer
    parenPos = InStr(CStr(cellValue), "(")
    
    If parenPos > 0 Then
        GetCompletedCount = Val(Left(CStr(cellValue), parenPos - 1))
    Else
        GetCompletedCount = Val(cellValue)
    End If
End Function

' --- Вспомогательная функция поиска ---
Private Function FindRowByText(ByVal textToFind As String) As Long
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsPlan = ThisWorkbook.Sheets("План")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row
    
    If lastRow < 2 Then Exit Function
    
    For i = 2 To lastRow
        ' ИЗМЕНЕНИЕ 3: Здесь тоже используем новую функцию, чтобы строки точно совпали
        Dim currentText As String
        currentText = CreateTaskDisplayString(wsPlan, i)
        
        If currentText = textToFind Then
            FindRowByText = i
            Exit Function
        End If
    Next i
End Function


' --- НОВАЯ ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ для чистого форматирования строки ---
Private Function CreateTaskDisplayString(ByVal ws As Worksheet, ByVal rowIdx As Long) As String
    ' Эта функция собирает строку для отображения, пропуская пустые ячейки.
    Dim parts As New Collection
    
    ' 1. Формируем первую часть (кол-во или *)
    Dim planValue As Variant: planValue = ws.Cells(rowIdx, COL_PLAN).value
    If CStr(planValue) = "*" Then
        parts.Add "*"
    Else
        Dim completedValue As Variant: completedValue = ws.Cells(rowIdx, COL_COMPLETED).value
        Dim numPlan As Long, numCompleted As Long
        If IsNumeric(planValue) Then numPlan = CLng(planValue)
        numCompleted = GetCompletedCount(completedValue)
        Dim remaining As Long: remaining = numPlan - numCompleted
        
        ' Пропускаем строку, если задача выполнена (на всякий случай)
        If remaining <= 0 Then Exit Function
        
        parts.Add remaining & "П"
    End If

    ' 2. Добавляем остальные части, только если они не пустые
    Dim material As String: material = Trim(CStr(ws.Cells(rowIdx, COL_MATERIAL).value))
    If material <> "" Then parts.Add material
    
    Dim kpz As String: kpz = Trim(CStr(ws.Cells(rowIdx, COL_KPZ).value))
    If kpz <> "" Then parts.Add kpz
    
    Dim customer As String: customer = Trim(CStr(ws.Cells(rowIdx, COL_CUSTOMER).value))
    If customer <> "" Then parts.Add customer
    
    Dim profile As String: profile = Trim(CStr(ws.Cells(rowIdx, COL_PROFILE).value))
    If profile <> "" Then parts.Add profile
    
    Dim color As String: color = Trim(CStr(ws.Cells(rowIdx, COL_COLOR).value))
    If color <> "" Then parts.Add color
    
    Dim lameli As String: lameli = Trim(CStr(ws.Cells(rowIdx, COL_LAMELI).value))
    If lameli <> "" Then parts.Add lameli & " шт."
    
    ' 3. Собираем все части в одну строку с помощью Join
    If parts.Count > 0 Then
        Dim arr() As String
        ReDim arr(1 To parts.Count)
        Dim i As Long
        For i = 1 To parts.Count
            arr(i) = parts(i)
        Next i
        CreateTaskDisplayString = Join(arr, " - ")
    End If
End Function


Private Sub UserForm_Zoom(Percent As Integer)

End Sub


