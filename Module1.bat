Option Explicit

' --- Модуль для создания листа "План" v4 (гибридный подход) ---

Sub CreatePlanSheet()
    Dim ws As Worksheet
    Dim sheetName As String

    sheetName = "План"

    Application.ScreenUpdating = False

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    Call CreateMainTable(ws)
    Call CreateTestData(ws)
    Call CreateReferenceLists(ws)
    
    
    Application.ScreenUpdating = True

    MsgBox "Лист '" & sheetName & "' с таблицей, списками и данными готов к работе!", vbInformation

End Sub

Sub CreateMainTable(ByRef ws As Worksheet)
    Dim tbl As ListObject
    Dim headers As Variant
    Dim i As Integer

    headers = Array( _
        "План (подвесы)", "Приоритет", "Выполнено (подвесы)", "Вид материалов", _
        "Годность подвеса", "Ответственный менеджер", "НОМЕР КПЗ", "Заказчик", _
        "Профиль", "Тип завески", "Вид обработки", "Толщина мкм", _
        "Цвет", "Подвесы, шт", "УСЛОВНЫЕ ПОДВЕСЫ, шт", "Ламели, шт" _
    )
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "tblPlan"
    tbl.TableStyle = "TableStyleMedium5"

    With tbl.HeaderRowRange
        .WrapText = True
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.color = vbBlack
    End With
    ws.Rows(1).EntireRow.AutoFit

    ' Отключаем перенос текста для всех строк данных, оставляя только в шапке
    ws.Range("A2:Q" & ws.Rows.Count).WrapText = False

    ws.Columns(1).ColumnWidth = 9.5:     ws.Columns(2).ColumnWidth = 3.5: ws.Columns(3).ColumnWidth = 11: ws.Columns(4).ColumnWidth = 16: ws.Columns(5).ColumnWidth = 19: ws.Columns(6).ColumnWidth = 15: ws.Columns(7).ColumnWidth = 11: ws.Columns(8).ColumnWidth = 24: ws.Columns(9).ColumnWidth = 35: ws.Columns(10).ColumnWidth = 10: ws.Columns(11).ColumnWidth = 12: ws.Columns(12).ColumnWidth = 8: ws.Columns(13).ColumnWidth = 15: ws.Columns(14).ColumnWidth = 9: ws.Columns(15).ColumnWidth = 6: ws.Columns(16).ColumnWidth = 9
End Sub
' --- 3. Создание справочных списков и привязка проверки данных (ОБЪЕДИНЕННЫЙ БЛОК) ---
Sub CreateReferenceLists(ByRef ws As Worksheet)
    Dim sourceColumnStart As Integer
    sourceColumnStart = 19 ' Начинаем с колонки S (19)

    ' --- Последовательно создаем и привязываем каждый справочный список ---
    ' Функция CreateAndLinkReferenceList теперь делает всю работу:
    ' создает таблицу, форматирует ее, настраивает проверку данных и возвращает новую стартовую колонку.

    sourceColumnStart = CreateAndLinkReferenceList(ws, "D", sourceColumnStart, _
        Array("Вид материалов", "Пороги", "Фасады", "Листы", "Детали", "Клюшки (растрав)", "Маркетинг (реклама, образцы)"), "TableStyleMedium2")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "E", sourceColumnStart, _
        Array("Годность подвеса", "Годный подвес", "Подбор цвета", "Монтаж на проволоку", "Брак (производство)", "Брак (поставщика)"), "TableStyleMedium3")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "F", sourceColumnStart, _
        Array("Ответственный менеджер", "Леонович Н.В."), "TableStyleMedium4")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "J", sourceColumnStart, _
        Array("Тип завески", "зажимы", "шпильки", "оснастка", "проволока"), "TableStyleMedium5")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "K", sourceColumnStart, _
        Array("Вид обработки", "Стандартная", "Допрастрав"), "TableStyleMedium6")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "L", sourceColumnStart, _
        Array("Толщина мкм", "10 мкм", "12 мкм", "15 мкм", "18 мкм", "20 мкм", "более 20 мкм", "25 мкм"), "TableStyleMedium7")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "M", sourceColumnStart, _
        Array("Цвет", "серебро", "золото", "бронза", "шампань", "черный (орг-ка)", "титан", "медь", "графит", "антрацит", "черный (эл/химия)", "согласно образца"), "TableStyleMedium8")
        
    ' --- Добавляем списки для Q и R ---
    ' Столбцы N, O, P пропускаются, поэтому продолжаем с той же логикой
    sourceColumnStart = CreateAndLinkReferenceList(ws, "Q", sourceColumnStart, _
        Array("Тест Q1", "Тест Q2", "Тест Q3"), "TableStyleMedium9")

    sourceColumnStart = CreateAndLinkReferenceList(ws, "R", sourceColumnStart, _
        Array("Тест R1", "Тест R2", "Тест R3", "Тест R4"), "TableStyleMedium10")
        
    ' Примечание: старая подпрограмма CreateReferenceList больше не нужна,
    ' так как вся её логика теперь находится внутри CreateAndLinkReferenceList.
End Sub
' --- 4. Универсальная функция для создания справочника и привязки его к столбцу (ОБЪЕДИНЕННЫЙ БЛОК) ---
Function CreateAndLinkReferenceList(ByRef ws As Worksheet, ByVal targetColumnLetter As String, ByVal sourceCol As Integer, ByVal values As Variant, ByVal tableStyleName As String) As Integer
    Dim i As Integer
    Dim sourceDataRange As Range
    Dim validationFormula As String
    Dim tbl As ListObject
    Dim tableName As String
    Dim targetColumn As Range

    ' --- 1. Записываем данные списка на лист и применяем форматирование ---
    ws.Columns(sourceCol).ColumnWidth = 15
    ws.Cells(1, sourceCol).Value = values(0)
    ws.Cells(1, sourceCol).Font.Bold = True
    
    For i = 1 To UBound(values)
        ws.Cells(i + 1, sourceCol).Value = values(i)
    Next i
    
    Set sourceDataRange = ws.Range(ws.Cells(1, sourceCol), ws.Cells(UBound(values) + 1, sourceCol))
    ws.Range(ws.Cells(2, sourceCol), ws.Cells(UBound(values) + 1, sourceCol)).Font.Bold = False

    ' --- 2. Создаем из данных таблицу Excel (ListObject) ---
    tableName = "tblRef" & Replace(values(0), " ", "") ' Уникальное имя для таблицы
    On Error Resume Next
    ws.ListObjects(tableName).Delete ' Удаляем, если таблица с таким именем уже есть
    On Error GoTo 0
    
    Set tbl = ws.ListObjects.Add(xlSrcRange, sourceDataRange, , xlYes)
    tbl.Name = tableName
    tbl.TableStyle = tableStyleName
    
    With tbl.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' --- 3. Динамически получаем адрес данных и создаем проверку (выпадающий список) ---
    If tbl.DataBodyRange Is Nothing Then
        ' Если в списке только заголовок (нет данных), то формула будет пустой
        ' Это предотвращает ошибку, хотя в вашем случае данные всегда есть
        validationFormula = ""
    Else
        ' Получаем адрес диапазона с данными из созданной таблицы. Это ключевой момент!
        validationFormula = "=" & tbl.DataBodyRange.Address
    End If
    
    If validationFormula <> "" Then
        Set targetColumn = ws.Columns(targetColumnLetter)
        With targetColumn.Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=validationFormula
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End If

    ' --- 4. Устанавливаем ширину разделительной колонки ---
    ws.Columns(sourceCol + 1).ColumnWidth = 1
    
    ' --- 5. Возвращаем номер следующей свободной колонки для следующего списка ---
    CreateAndLinkReferenceList = sourceCol + 2
End Function
Sub CreateTestData(ByRef ws As Worksheet)
    Dim tbl As ListObject, newRow As ListRow
    Set tbl = ws.ListObjects("tblPlan")
    If tbl.ListRows.Count > 0 Then tbl.DataBodyRange.Delete


    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(50, 1, 10, "Пороги", "Годный подвес", "Леонович Н.В.", 140, "РП", "ЮП-2076", "зажимы", "стандартная", "15 мкм", "черный (орг-ка)", 1, "", 22)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(100, 1, 0, "Пороги", "Годный подвес", "", 138, "РП", "ЮП-1875", "зажимы", "стандартная", "15 мкм", "серебро", 1, "", 36)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(20, "", 5, "Детали", "Годный подвес", "Леонович Н.В.", 6761, "Горкомплекс", "Урны", "оснастка", "допрастрав", "15 мкм", "с31", 1, "", 8)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(15, "", 15, "Листы", "Годный подвес", "Леонович Н.В.", 6765, "Минский_авиаремонтный", "1.2x1500x3000", "зажимы", "стандартная", "15 мкм", "серебро", 1, "", 1)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(30, "", 0, "Фасады", "Подбор цвета", "", 111, "Заказчик А", "Профиль Z", "шпильки", "стандартная", "20 мкм", "золото", 1, "", 15)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(25, "", 0, "Детали", "Монтаж на проволоку", "", 222, "Заказчик Б", "Профиль Y", "проволока", "стандартная", "18 мкм", "бронза", 1, "", 10)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array("*", "", "", "Клюшки (растрав)", "", "", "", "", "", "", "", "", "", 1, "", "")
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(5, 1, 0, "Маркетинг (реклама, образцы)", "Годный подвес", "Леонович Н.В.", 333, "Отдел маркетинга", "Образец 123", "зажимы", "стандартная", "12 мкм", "шампань", 1, "", 5)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(40, "", 20, "Пороги", "Брак (производство)", "", 117, "РП", "ЮП-3094", "зажимы", "допрастрав", "15 мкм", "серебро", 1, "", 47)
    Set newRow = tbl.ListRows.Add: newRow.Range.Value = Array(60, "", 0, "Листы", "Годный подвес", "Леонович Н.В.", 444, "Заказчик В", "2.0x1000x2000", "шпильки", "стандартная", "25 мкм", "титан", 1, "", 5)

If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Font.Bold = False
    End If

End Sub

