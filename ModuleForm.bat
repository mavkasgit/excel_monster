Option Explicit

' --- КОД С РАЗДЕЛЕНИЕМ СПИСКОВ НА "В РАБОТЕ" И "ОЖИДАЮТ" ---

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

' === Переменные модуля ===
' Для отслеживания, какой из списков активен в данный момент
Private m_ActiveListBox As Object ' MSForms.ListBox

' --- События формы ---

Private Sub UserForm_Initialize()
    ' --- Логика загрузки состояния чекбокса ---
    Dim recordTimeSetting As Name
    On Error Resume Next
    Set recordTimeSetting = ThisWorkbook.Names("RecordTimeSetting")
    On Error GoTo 0

    If recordTimeSetting Is Nothing Then
        ThisWorkbook.Names.Add Name:="RecordTimeSetting", RefersTo:="=TRUE"
        Me.chkRecordTime.value = True
    Else
        Me.chkRecordTime.value = Evaluate(recordTimeSetting.RefersTo)
    End If

    ' По умолчанию считаем активным верхний список
    Set m_ActiveListBox = Me.lstTasks

    ' --- Основная логика ---
    Call PopulateList
End Sub

Private Sub chkRecordTime_Click()
    ' Сохраняем состояние чекбокса
    On Error Resume Next
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

' --- НОВЫЕ ОБРАБОТЧИКИ КЛИКОВ ДЛЯ ДВУХ СПИСКОВ ---
Private Sub lstInProgressTasks_Click()
    ' Этот список стал активным
    Set m_ActiveListBox = Me.lstInProgressTasks
    ' Сбрасываем выделение в другом списке
    Me.lstTasks.ListIndex = -1
End Sub

Private Sub lstTasks_Click()
    ' Этот список стал активным
    Set m_ActiveListBox = Me.lstTasks
    ' Сбрасываем выделение в другом списке
    Me.lstInProgressTasks.ListIndex = -1
End Sub


' --- Логика обработки клика (теперь работает с активным списком) ---
Private Sub HandleTask(CloseAfter As Boolean)
    ' Проверяем, есть ли активный список и выбрана ли в нем задача
    If m_ActiveListBox Is Nothing Then Exit Sub
    If m_ActiveListBox.ListIndex = -1 Then Exit Sub
    
    Dim selectedText As String
    selectedText = m_ActiveListBox.value ' Берем значение из активного списка

    If Left(selectedText, 3) = "---" Then Exit Sub
    
    ' --- УЛУЧШЕННАЯ ЛОГИКА С ЯКОРЕМ ---
    Dim stableIdentifier As String
    Dim separatorPos As Long
    separatorPos = InStr(selectedText, " - ")
    
    If separatorPos > 0 Then
        stableIdentifier = Mid(selectedText, separatorPos)
    Else
        stableIdentifier = selectedText
    End If

    Dim initialRowIndex As Long
    initialRowIndex = FindRowByText(selectedText)
    
    If initialRowIndex > 0 Then
        Call ExecuteTask(initialRowIndex)
        
        If CloseAfter Then
            Unload Me
            Exit Sub
        End If

        ' Перезагружаем оба списка
        Call PopulateList

        ' ИЩЕМ ЗАДАЧУ В ОБОИХ СПИСКАХ ПО СТАБИЛЬНОМУ ИДЕНТИФИКАТОРУ
        Dim i As Long
        Dim isFound As Boolean
        isFound = False
        
        ' Сначала ищем в списке "В работе"
        For i = 0 To Me.lstInProgressTasks.ListCount - 1
            If InStr(CStr(Me.lstInProgressTasks.List(i)), stableIdentifier) > 0 Then
                Me.lstInProgressTasks.ListIndex = i
                Set m_ActiveListBox = Me.lstInProgressTasks ' Делаем его активным
                Me.lstTasks.ListIndex = -1 ' Сбрасываем другой
                isFound = True
                Exit For
            End If
        Next i
        
        ' Если не нашли, ищем в списке "Ожидают"
        If Not isFound Then
            For i = 0 To Me.lstTasks.ListCount - 1
                If InStr(CStr(Me.lstTasks.List(i)), stableIdentifier) > 0 Then
                    Me.lstTasks.ListIndex = i
                    Set m_ActiveListBox = Me.lstTasks ' Делаем его активным
                    Me.lstInProgressTasks.ListIndex = -1 ' Сбрасываем другой
                    isFound = True
                    Exit For
                End If
            Next i
        End If
        
        If Not isFound Then
            ' Задача, видимо, завершена
        End If
        
        m_ActiveListBox.SetFocus
        
    Else
        MsgBox "Не удалось найти исходную строку на листе: '" & selectedText & "'. Список будет обновлен.", vbExclamation
        Call PopulateList
    End If
End Sub

' --- Основная процедура отображения (теперь для двух списков) ---
Public Sub PopulateList()
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Три коллекции для трех типов задач
    Dim inProgressTasks As New Collection
    Dim priorityTasks As New Collection
    Dim standardTasks As New Collection
    Dim totalHangersSum As Double
    
    On Error GoTo 0

    Set wsPlan = ThisWorkbook.Sheets("План")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row

    ' Очищаем оба списка
    Me.lstTasks.Clear
    Me.lstInProgressTasks.Clear
    Me.lstTasks.ColumnCount = 1
    Me.lstTasks.ColumnWidths = "700"
    Me.lstInProgressTasks.ColumnCount = 1
    Me.lstInProgressTasks.ColumnWidths = "700"
    
    If lastRow < 2 Then
        Me.lstTasks.AddItem "На листе 'План' нет данных."
        Exit Sub
    End If
    
    For i = 2 To lastRow
        Dim planValue As Variant: planValue = wsPlan.Cells(i, COL_PLAN).value
        Dim completedValue As Variant: completedValue = wsPlan.Cells(i, COL_COMPLETED).value
        
        Dim numPlan As Long, numCompleted As Long
        If CStr(planValue) = "*" Then
            numPlan = 999999 ' Условное большое число для задач "*"
        Else
            If IsNumeric(planValue) Then numPlan = CLng(planValue)
        End If
        numCompleted = GetCompletedCount(completedValue)
        
        Dim remaining As Long
        remaining = numPlan - numCompleted
        
        If remaining > 0 Then
            If CStr(planValue) <> "*" Then
                totalHangersSum = totalHangersSum + remaining
            End If
            
            Dim displayText As String
            displayText = CreateTaskDisplayString(wsPlan, i)
            
            ' --- НОВАЯ ЛОГИКА РАСПРЕДЕЛЕНИЯ ---
            If numCompleted > 0 Then
                ' Если задача уже начата, она идет в список "В работе"
                inProgressTasks.Add displayText
            Else
                ' Иначе, распределяем по приоритету
                Dim isPriority As Boolean
                Dim priorityCheckValue As Variant
                priorityCheckValue = wsPlan.Cells(i, COL_PRIORITY).value
                
                isPriority = False
                If IsNumeric(priorityCheckValue) Then
                    If CDbl(priorityCheckValue) > 0 Then isPriority = True
                End If
                
                If isPriority Then
                    priorityTasks.Add displayText
                Else
                    standardTasks.Add displayText
                End If
            End If
        End If
    Next i

    Dim item As Variant
    
    ' --- Заполняем список "В работе" ---
    If inProgressTasks.Count > 0 Then
        For Each item In inProgressTasks
            Me.lstInProgressTasks.AddItem item
        Next item
    Else
        Me.lstInProgressTasks.AddItem "Нет начатых задач."
    End If

    ' --- Заполняем список "Ожидают" ---
    If priorityTasks.Count > 0 Or standardTasks.Count > 0 Then
        For Each item In priorityTasks
            Me.lstTasks.AddItem item
        Next item
        
        If priorityTasks.Count > 0 And standardTasks.Count > 0 Then
            Me.lstTasks.AddItem "--- (обычные) ---"
        End If
        
        For Each item In standardTasks
            Me.lstTasks.AddItem item
        Next item
    Else
         Me.lstTasks.AddItem "Нет задач в ожидании."
    End If


    Dim totalTasks As Long
    totalTasks = inProgressTasks.Count + priorityTasks.Count + standardTasks.Count
    
    If totalTasks = 0 Then
        Me.Controls("lblTotal").Caption = "Общее кол-во подвесов: 0"
    Else
        Me.Controls("lblTotal").Caption = "Общее кол-во подвесов: " & totalHangersSum
    End If
End Sub

' --- Вспомогательные функции (без изменений) ---
Private Function GetCompletedCount(ByVal cellValue As Variant) As Long
    Dim parenPos As Integer
    parenPos = InStr(CStr(cellValue), "(")
    
    If parenPos > 0 Then
        GetCompletedCount = Val(Left(CStr(cellValue), parenPos - 1))
    Else
        GetCompletedCount = Val(cellValue)
    End If
End Function

Private Function FindRowByText(ByVal textToFind As String) As Long
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsPlan = ThisWorkbook.Sheets("План")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row
    
    If lastRow < 2 Then Exit Function
    
    For i = 2 To lastRow
        Dim currentText As String
        currentText = CreateTaskDisplayString(wsPlan, i)
        
        If currentText = textToFind Then
            FindRowByText = i
            Exit Function
        End If
    Next i
End Function

Private Function CreateTaskDisplayString(ByVal ws As Worksheet, ByVal rowIdx As Long) As String
    Dim parts As New Collection
    
    Dim planValue As Variant: planValue = ws.Cells(rowIdx, COL_PLAN).value
    If CStr(planValue) = "*" Then
        parts.Add "*"
    Else
        Dim completedValue As Variant: completedValue = ws.Cells(rowIdx, COL_COMPLETED).value
        Dim numPlan As Long, numCompleted As Long
        If IsNumeric(planValue) Then numPlan = CLng(planValue)
        numCompleted = GetCompletedCount(completedValue)
        Dim remaining As Long: remaining = numPlan - numCompleted
        
        If remaining <= 0 Then Exit Function
        
        parts.Add remaining & "П"
    End If

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