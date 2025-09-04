Option Explicit

' --- ÊÎÄ Ñ ÓËÓ×ØÅÍÈßÌÈ: ÑÓÌÌÀ ÏÎÄÂÅÑÎÂ È ×ÈÑÒÎÅ ÔÎÐÌÀÒÈÐÎÂÀÍÈÅ ---

' === Êîíñòàíòû ===
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

' --- Ñîáûòèÿ ---
Private Sub UserForm_Initialize()
    ' --- Ëîãèêà çàãðóçêè ñîñòîÿíèÿ ÷åêáîêñà èç Èìåíîâàííîãî Äèàïàçîíà ---
    Dim recordTimeSetting As Name
    
    On Error Resume Next
    Set recordTimeSetting = ThisWorkbook.Names("RecordTimeSetting")
    On Error GoTo 0

    If recordTimeSetting Is Nothing Then
        ' Èìÿ íå ñóùåñòâóåò, ñîçäàåì åãî ñî çíà÷åíèåì ïî óìîë÷àíèþ TRUE
        ThisWorkbook.Names.Add Name:="RecordTimeSetting", RefersTo:="=TRUE"
        Me.chkRecordTime.value = True
    Else
        ' Èìÿ ñóùåñòâóåò, ñ÷èòûâàåì åãî çíà÷åíèå
        ' RefersTo âîçâðàùàåò ñòðîêó "=TRUE" èëè "=FALSE", Evaluate ïðåîáðàçóåò åå â ëîãè÷åñêîå çíà÷åíèå
        Me.chkRecordTime.value = Evaluate(recordTimeSetting.RefersTo)
    End If

    ' --- Îñíîâíàÿ ëîãèêà ---
    Call PopulateList
End Sub

Private Sub chkRecordTime_Click()
    ' Ñîõðàíÿåì ñîñòîÿíèå ÷åêáîêñà â Èìåíîâàííûé Äèàïàçîí
    On Error Resume Next ' Íà ñëó÷àé åñëè èìÿ áûëî óäàëåíî
    ThisWorkbook.Names("RecordTimeSetting").RefersTo = "=" & UCase(CStr(Me.chkRecordTime.value))
    On Error GoTo 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' --- Êíîïêè ---
Private Sub cmdAddTask_Click()
    Call HandleTask(CloseAfter:=False)
End Sub
Private Sub cmdAddAndClose_Click()
    Call HandleTask(CloseAfter:=True)
End Sub

' --- Ëîãèêà îáðàáîòêè êëèêà ---
Private Sub HandleTask(CloseAfter As Boolean)
    If Me.lstTasks.ListIndex = -1 Then Exit Sub
    
    Dim selectedText As String
    selectedText = Me.lstTasks.value ' Çàïîìèíàåì òåêñò ÄÎ âñåõ èçìåíåíèé

    If Left(selectedText, 3) = "---" Then Exit Sub
    
    ' --- ÍÎÂÀß ÓËÓ×ØÅÍÍÀß ËÎÃÈÊÀ ---
    
    ' 1. Èçâëåêàåì "ñòàáèëüíóþ" ÷àñòü òåêñòà çàäà÷è, êîòîðàÿ íå ìåíÿåòñÿ.
    '    Ýòî áóäåò íàø "ÿêîðü" äëÿ ïîèñêà ïîñëå âñåõ ïåðåìåùåíèé.
    Dim stableIdentifier As String
    Dim separatorPos As Long
    separatorPos = InStr(selectedText, " - ")
    
    If separatorPos > 0 Then
        stableIdentifier = Mid(selectedText, separatorPos) ' Ïîëó÷àåì " - KPZ123 - Çàêàç÷èê..."
    Else
        stableIdentifier = selectedText ' Íà ñëó÷àé, åñëè â ñòðîêå íåò äåôèñà (íàïðèìåð, òîëüêî "*")
    End If

    ' 2. Íàõîäèì ïåðâîíà÷àëüíûé íîìåð ñòðîêè, ÷òîáû ïåðåäàòü åãî â ExecuteTask
    Dim initialRowIndex As Long
    initialRowIndex = FindRowByText(selectedText)
    
    If initialRowIndex > 0 Then
        ' 3. Âûïîëíÿåì çàäà÷ó. Â ÝÒÎÒ ÌÎÌÅÍÒ ÑÒÐÎÊÀ ÏÅÐÅÌÅÙÀÅÒÑß,
        '    è initialRowIndex ñòàíîâèòñÿ íåàêòóàëüíûì!
        Call ExecuteTask(initialRowIndex)
        
        If CloseAfter Then
            Unload Me
            Exit Sub
        End If

        ' Ïåðåçàãðóæàåì ñïèñîê ñ ó÷åòîì âñåõ èçìåíåíèé íà ëèñòå
        Call PopulateList

        ' 4. ÈÙÅÌ ÇÀÄÀ×Ó Â ÍÎÂÎÌ ÑÏÈÑÊÅ ÏÎ ÑÒÀÁÈËÜÍÎÌÓ ÈÄÅÍÒÈÔÈÊÀÒÎÐÓ
        Dim i As Long
        Dim isFound As Boolean
        isFound = False
        
        For i = 0 To Me.lstTasks.ListCount - 1
            ' Ìû èùåì ñòðîêó, êîòîðàÿ ÇÀÊÀÍ×ÈÂÀÅÒÑß íà íàø "ÿêîðü".
            ' Ýòî ñðàáîòàåò, äàæå åñëè êîëè÷åñòâî â íà÷àëå èçìåíèëîñü (ñ 10Ï íà 9Ï).
            ' Èñïîëüçóåì CStr äëÿ íàäåæíîñòè, åñëè â ñïèñêå îêàæåòñÿ íå òåêñòîâîå çíà÷åíèå.
            If InStr(CStr(Me.lstTasks.List(i)), stableIdentifier) > 0 Then
                Me.lstTasks.ListIndex = i ' Íàøëè! Âûäåëÿåì.
                isFound = True
                Exit For ' Âûõîäèì, ò.ê. çàäà÷à íàéäåíà
            End If
        Next i
        
        ' 5. Åñëè ïîñëå öèêëà çàäà÷à íå íàéäåíà, çíà÷èò, îíà áûëà ïîñëåäíåé è çàâåðøèëàñü.
        If Not isFound Then
            MsgBox "Çàäà÷à '" & selectedText & "' ïîëíîñòüþ âûïîëíåíà è óáðàíà èç ñïèñêà.", vbInformation, "Çàäà÷à çàâåðøåíà"
            ' Îïöèîíàëüíî: âûäåëèòü ïåðâóþ çàäà÷ó â ñïèñêå, ÷òîáû ôîêóñ íå ïðîïàäàë
            If Me.lstTasks.ListCount > 0 Then Me.lstTasks.ListIndex = 0
        End If
        
        Me.lstTasks.SetFocus
        
    Else
        ' Ýòà îøèáêà ìîæåò âîçíèêíóòü, åñëè äàííûå íà ëèñòå èçìåíèëèñü ìåæäó PopulateList è êëèêîì
        MsgBox "Íå óäàëîñü íàéòè èñõîäíóþ ñòðîêó íà ëèñòå: '" & selectedText & "'. Ñïèñîê áóäåò îáíîâëåí.", vbExclamation
        Call PopulateList
    End If
End Sub
' --- Îñíîâíàÿ ïðîöåäóðà îòîáðàæåíèÿ ---
Public Sub PopulateList()
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Dim priorityTasks As New Collection
    Dim standardTasks As New Collection
    Dim totalHangersSum As Double ' ÈÇÌÅÍÅÍÈÅ 2: Ïåðåìåííàÿ äëÿ ñóììû ïîäâåñîâ
    
    On Error GoTo 0

    Set wsPlan = ThisWorkbook.Sheets("Ïëàí")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row

    Me.lstTasks.Clear
    Me.lstTasks.ColumnCount = 1
    Me.lstTasks.ColumnWidths = "700"
    
    If lastRow < 2 Then
        Me.lstTasks.AddItem "Íà ëèñòå 'Ïëàí' íåò äàííûõ."
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
            ' ÈÇÌÅÍÅÍÈÅ 2: Ñóììèðóåì ïîäâåñû òîëüêî äëÿ ÷èñëîâûõ çàäà÷
            If CStr(planValue) <> "*" Then
                totalHangersSum = totalHangersSum + remaining
            End If
            
            ' ÈÇÌÅÍÅÍÈÅ 3: Èñïîëüçóåì íîâóþ ôóíêöèþ äëÿ ÷èñòîãî ôîðìàòèðîâàíèÿ ñòðîêè
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
    
    ' ÈÇÌÅÍÅÍÈÅ 1: Ýòà ëîãèêà óæå âåðíà. Ðàçäåëèòåëü ïîÿâèòñÿ, òîëüêî åñëè åñòü ÎÁÀ òèïà çàäà÷.
    If priorityTasks.Count > 0 And standardTasks.Count > 0 Then
        Me.lstTasks.AddItem "--- ÑÐÎ×ÍÎ --- ÑÐÎ×ÍÎ --- ÑÐÎ×ÍÎ --- ÑÐÎ×ÍÎ ---"
    End If
    
    For Each item In standardTasks
        Me.lstTasks.AddItem item
    Next item

    Dim totalTasks As Long
    totalTasks = priorityTasks.Count + standardTasks.Count
    
    If totalTasks = 0 Then
        Me.lstTasks.AddItem "Âñå çàäà÷è âûïîëíåíû."
        Me.Controls("lblTotal").Caption = "Îáùåå êîë-âî ïîäâåñîâ: 0"
    Else
        ' ÈÇÌÅÍÅÍÈÅ 2: Îáíîâëÿåì íàäïèñü, ÷òîáû ïîêàçûâàòü ñóììó ïîäâåñîâ
        Me.Controls("lblTotal").Caption = "Îáùåå êîë-âî ïîäâåñîâ: " & totalHangersSum
    End If
End Sub

' --- Âñïîìîãàòåëüíàÿ ôóíêöèÿ äëÿ ïàðñèíãà êîëè÷åñòâà âûïîëíåííûõ ---
Private Function GetCompletedCount(ByVal cellValue As Variant) As Long
    Dim parenPos As Integer
    parenPos = InStr(CStr(cellValue), "(")
    
    If parenPos > 0 Then
        GetCompletedCount = Val(Left(CStr(cellValue), parenPos - 1))
    Else
        GetCompletedCount = Val(cellValue)
    End If
End Function

' --- Âñïîìîãàòåëüíàÿ ôóíêöèÿ ïîèñêà ---
Private Function FindRowByText(ByVal textToFind As String) As Long
    Dim wsPlan As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set wsPlan = ThisWorkbook.Sheets("Ïëàí")
    lastRow = wsPlan.Cells(wsPlan.Rows.Count, COL_PLAN).End(xlUp).Row
    
    If lastRow < 2 Then Exit Function
    
    For i = 2 To lastRow
        ' ÈÇÌÅÍÅÍÈÅ 3: Çäåñü òîæå èñïîëüçóåì íîâóþ ôóíêöèþ, ÷òîáû ñòðîêè òî÷íî ñîâïàëè
        Dim currentText As String
        currentText = CreateTaskDisplayString(wsPlan, i)
        
        If currentText = textToFind Then
            FindRowByText = i
            Exit Function
        End If
    Next i
End Function


' --- ÍÎÂÀß ÂÑÏÎÌÎÃÀÒÅËÜÍÀß ÔÓÍÊÖÈß äëÿ ÷èñòîãî ôîðìàòèðîâàíèÿ ñòðîêè ---
Private Function CreateTaskDisplayString(ByVal ws As Worksheet, ByVal rowIdx As Long) As String
    ' Ýòà ôóíêöèÿ ñîáèðàåò ñòðîêó äëÿ îòîáðàæåíèÿ, ïðîïóñêàÿ ïóñòûå ÿ÷åéêè.
    Dim parts As New Collection
    
    ' 1. Ôîðìèðóåì ïåðâóþ ÷àñòü (êîë-âî èëè *)
    Dim planValue As Variant: planValue = ws.Cells(rowIdx, COL_PLAN).value
    If CStr(planValue) = "*" Then
        parts.Add "*"
    Else
        Dim completedValue As Variant: completedValue = ws.Cells(rowIdx, COL_COMPLETED).value
        Dim numPlan As Long, numCompleted As Long
        If IsNumeric(planValue) Then numPlan = CLng(planValue)
        numCompleted = GetCompletedCount(completedValue)
        Dim remaining As Long: remaining = numPlan - numCompleted
        
        ' Ïðîïóñêàåì ñòðîêó, åñëè çàäà÷à âûïîëíåíà (íà âñÿêèé ñëó÷àé)
        If remaining <= 0 Then Exit Function
        
        parts.Add remaining & "Ï"
    End If

    ' 2. Äîáàâëÿåì îñòàëüíûå ÷àñòè, òîëüêî åñëè îíè íå ïóñòûå
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
    If lameli <> "" Then parts.Add lameli & " øò."
    
    ' 3. Ñîáèðàåì âñå ÷àñòè â îäíó ñòðîêó ñ ïîìîùüþ Join
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


