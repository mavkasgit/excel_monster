Option Explicit

Private IsAdvancedBackupRunning As Boolean ' ИЗМЕНЕНО: Теперь это внутренняя (Private) переменная модуля.
Public NextCheckTime As Date
Private Const BaseBackupPath As String = "D:\VibeCoding\Backups"

'========================================================================================
'= УПРАВЛЕНИЕ АВТОМАТИЧЕСКИМ ЦИКЛОМ
'========================================================================================

' ИНИЦИАЛИЗИРУЕТ СИСТЕМУ (вызывается из ThisWorkbook при открытии файла)
Public Sub InitializeBackupSystem()
    If IsAdvancedBackupRunning Then Exit Sub
    IsAdvancedBackupRunning = True
    CreateBackupFolders
    Application.StatusBar = "Система авто-бэкапов активна. Проверка каждые 5 минут."
    AdvancedBackupWorker
End Sub

' НОВАЯ ПРОЦЕДУРА для корректной остановки (вызывается из ThisWorkbook при закрытии файла)
Public Sub ShutdownBackupSystem()
    If Not IsAdvancedBackupRunning Then Exit Sub
    IsAdvancedBackupRunning = False
    On Error Resume Next
    Application.OnTime EarliestTime:=NextCheckTime, Procedure:="AdvancedBackupWorker", Schedule:=False
    On Error GoTo 0
    Application.StatusBar = False
End Sub

' ФОНОВЫЙ ПЛАНИРОВЩИК (просыпается каждые 5 минут)
Public Sub AdvancedBackupWorker()
    If Not IsAdvancedBackupRunning Then Exit Sub
    PerformBackupLogic
    If IsAdvancedBackupRunning Then
        NextCheckTime = Now + TimeValue("00:05:00")
        Application.OnTime NextCheckTime, "AdvancedBackupWorker"
    End If
End Sub

'========================================================================================
'= ОСНОВНАЯ ЛОГИКА РЕЗЕРВНОГО КОПИРОВАНИЯ (СЕРДЦЕ МАКРОСА)
'========================================================================================

Private Sub PerformBackupLogic()
    Application.ScreenUpdating = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- 1. ЧАСОВЫЕ БЭКАПЫ ---
    Dim hourlyPath As String: hourlyPath = BaseBackupPath & "\Hourly"
    If DateDiff("h", GetLatestFileDate(hourlyPath), Now) >= 1 Then
        Dim hourlyFileName As String
        hourlyFileName = fso.BuildPath(hourlyPath, "backup_" & Format(Now, "YYYY-MM-DD_HH-00") & ".xlsm")
        ThisWorkbook.SaveCopyAs hourlyFileName
        CleanupOldFiles hourlyPath, 24
    End If

    ' --- 2. ДНЕВНЫЕ БЭКАПЫ ---
    Dim dailyPath As String: dailyPath = BaseBackupPath & "\Daily"
    If Int(GetLatestFileDate(dailyPath)) <> Int(Now) Then
        Dim dailyFileName As String
        dailyFileName = fso.BuildPath(dailyPath, "backup_" & Format(Now, "dddd") & ".xlsm")
        ThisWorkbook.SaveCopyAs dailyFileName
    End If

    ' --- 3. МЕСЯЧНЫЕ БЭКАПЫ ---
    Dim monthlyPath As String: monthlyPath = BaseBackupPath & "\Monthly"
    If Format(GetLatestFileDate(monthlyPath), "YYYY-MM") <> Format(Now, "YYYY-MM") Then
        Dim monthlyFileName As String
        monthlyFileName = fso.BuildPath(monthlyPath, "backup_" & Format(Now, "YYYY-MMMM") & ".xlsm")
        ThisWorkbook.SaveCopyAs monthlyFileName
    End If

    Set fso = Nothing
    Application.ScreenUpdating = True
End Sub

'========================================================================================
'= ДИАГНОСТИКА И РУЧНОЙ ЗАПУСК
'========================================================================================

Public Sub GenerateFullBackupReport()
    Dim reportString As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    reportString = "--- АНАЛИЗ СОСТОЯНИЯ БЭКАПОВ ---" & vbCrLf & "Текущее системное время: " & Now & vbCrLf & "=======================================" & vbCrLf & vbCrLf
    Dim hourlyPath As String: hourlyPath = BaseBackupPath & "\Hourly"
    Dim lastHourly As Date: lastHourly = GetLatestFileDate(hourlyPath)
    reportString = reportString & "АНАЛИЗ: Часовые бэкапы" & vbCrLf & "----------------------------------------" & vbCrLf
    If lastHourly > 0 Then reportString = reportString & "Время последнего бэкапа: " & lastHourly & vbCrLf Else reportString = reportString & "Время последнего бэкапа: (не найдено)" & vbCrLf
    Dim hoursDiff As Long: hoursDiff = DateDiff("h", lastHourly, Now)
    reportString = reportString & "Прошло полных часов: " & hoursDiff & " (Требуется >= 1)" & vbCrLf
    If hoursDiff >= 1 Then reportString = reportString & "РЕКОМЕНДАЦИЯ: **Создать новый часовой бэкап.**" & vbCrLf Else reportString = reportString & "РЕКОМЕНДАЦИЯ: Действий не требуется." & vbCrLf
    If fso.FolderExists(hourlyPath) Then reportString = reportString & "Файлов в папке: " & fso.GetFolder(hourlyPath).Files.Count & " (Лимит: 24)" & vbCrLf
    reportString = reportString & vbCrLf
    Dim dailyPath As String: dailyPath = BaseBackupPath & "\Daily"
    Dim lastDaily As Date: lastDaily = GetLatestFileDate(dailyPath)
    reportString = reportString & "АНАЛИЗ: Дневные бэкапы" & vbCrLf & "----------------------------------------" & vbCrLf
    If lastDaily > 0 Then reportString = reportString & "Дата последнего бэкапа: " & Format(lastDaily, "Long Date") & vbCrLf Else reportString = reportString & "Дата последнего бэкапа: (не найдено)" & vbCrLf
    If Int(lastDaily) <> Int(Now) Then reportString = reportString & "РЕКОМЕНДАЦИЯ: **Создать новый дневной бэкап.**" & vbCrLf Else reportString = reportString & "РЕКОМЕНДАЦИЯ: Действий не требуется." & vbCrLf
    reportString = reportString & vbCrLf
    Dim monthlyPath As String: monthlyPath = BaseBackupPath & "\Monthly"
    Dim lastMonthly As Date: lastMonthly = GetLatestFileDate(monthlyPath)
    reportString = reportString & "АНАЛИЗ: Месячные бэкапы" & vbCrLf & "----------------------------------------" & vbCrLf
    If lastMonthly > 0 Then reportString = reportString & "Месяц последнего бэкапа: " & Format(lastMonthly, "MMMM YYYY") & vbCrLf Else reportString = reportString & "Месяц последнего бэкапа: (не найдено)" & vbCrLf
    If Format(lastMonthly, "YYYY-MM") <> Format(Now, "YYYY-MM") Then reportString = reportString & "РЕКОМЕНДАЦИЯ: **Создать новый месячный бэкап.**" & vbCrLf Else reportString = reportString & "РЕКОМЕНДАЦИЯ: Действий не требуется." & vbCrLf
    reportString = reportString & vbCrLf & "=======================================" & vbCrLf & "Выполнить рекомендуемые действия сейчас?"
    Dim userResponse As VbMsgBoxResult
    userResponse = MsgBox(reportString, vbQuestion + vbYesNo, "Анализ и подтверждение")
    If userResponse = vbYes Then
        PerformBackupLogic
        MsgBox "Операция выполнена. Резервные копии созданы/обновлены согласно отчету.", vbInformation, "Завершено"
    Else
        MsgBox "Операция отменена пользователем.", vbInformation, "Отмена"
    End If
End Sub

'========================================================================================
'= ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
'========================================================================================

Private Function GetLatestFileDate(folderPath As String) As Date
    Dim fso As Object, folder As Object, file As Object
    Dim latestDate As Date
    Set fso = CreateObject("Scripting.FileSystemObject")
    latestDate = 0
    If fso.FolderExists(folderPath) Then
        Set folder = fso.GetFolder(folderPath)
        For Each file In folder.Files
            If file.DateLastModified > latestDate Then latestDate = file.DateLastModified
        Next
    End If
    GetLatestFileDate = latestDate
End Function

Private Sub CleanupOldFiles(folderPath As String, maxFiles As Long)
    ' Удаляет самые старые файлы в папке, если их общее количество превышает maxFiles.
    ' Эта логика основана на КОЛИЧЕСТВЕ файлов, сохраняя 24 самых новых.
    Dim fso As Object, folder As Object, file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Sub
    
    Set folder = fso.GetFolder(folderPath)
    
    ' Если количество файлов не превышает лимит, никаких действий не требуется.
    If folder.Files.Count <= maxFiles Then Exit Sub
    
    ' Используем SortedList для автоматической сортировки файлов по дате создания.
    Dim sortedFiles As Object
    Set sortedFiles = CreateObject("System.Collections.SortedList")
    
    ' 1. Собираем все файлы в отсортированный список.
    For Each file In folder.Files
        ' Ключом является дата модификации файла в формате, удобном для сортировки,
        ' с добавлением полного пути для гарантии уникальности ключа.
        Dim sortKey As String
        sortKey = Format(file.DateLastModified, "YYYYMMDDHHNNSS") & "_" & file.Path
        
        ' Значением является полный путь к файлу.
        If Not sortedFiles.ContainsKey(sortKey) Then
            sortedFiles.Add sortKey, file.Path
        End If
    Next
    
    ' 2. Определяем, сколько самых старых файлов нужно удалить.
    Dim deleteCount As Long
    deleteCount = sortedFiles.Count - maxFiles
    
    ' 3. Удаляем необходимое количество самых старых файлов.
    ' Так как список отсортирован по возрастанию даты, самые старые файлы находятся в начале.
    Dim i As Long
    For i = 0 To deleteCount - 1
        Dim filePathToDelete As String
        filePathToDelete = sortedFiles.GetByIndex(i)
        
        On Error Resume Next ' Игнорируем ошибки при удалении одного файла
        fso.DeleteFile filePathToDelete, True
        On Error GoTo 0
    Next
    
    Set fso = Nothing
    Set sortedFiles = Nothing
End Sub

Private Sub CreateBackupFolders()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(BaseBackupPath) Then fso.CreateFolder BaseBackupPath
    If Not fso.FolderExists(BaseBackupPath & "\Hourly") Then fso.CreateFolder BaseBackupPath & "\Hourly"
    If Not fso.FolderExists(BaseBackupPath & "\Daily") Then fso.CreateFolder BaseBackupPath & "\Daily"
    If Not fso.FolderExists(BaseBackupPath & "\Monthly") Then fso.CreateFolder BaseBackupPath & "\Monthly"
    Set fso = Nothing
End Sub