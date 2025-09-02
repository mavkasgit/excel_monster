<?xml version="1.0" encoding="UTF-8"?>
<VBA_Rules>
    <General>
        <Rule>Код должен быть рабочим в VBA Editor (Alt+F11), без сторонних библиотек, кроме стандартных Excel-объектов.</Rule>
        <Rule>Комментарии обязательны: пояснять логику кода и ключевые части (циклы, условия, работу с диапазонами).</Rule>
        <Rule>Код должен быть структурирован: отступы, аккуратное оформление, объявления переменных в начале процедур.</Rule>
        <Rule>Всегда использовать Option Explicit для избежания ошибок из-за неявных переменных.</Rule>
        <Rule>Функции всегда объявлять как Public и вызывать напрямую, без обращения через имя модуля.</Rule>
        <Rule>Код должен быть разделён на части: каждая процедура выполняет одну задачу для упрощения отладки.</Rule>
    </General>

    <CommonErrors>
        <Error>Неправильные ссылки на диапазоны. Нельзя использовать Range("A1") без указания листа.</Error>
        <Fix>Использовать Worksheets("Лист1").Range("A1")</Fix>

        <Error>Ошибки при циклах: неверное определение lastRow.</Error>
        <Fix>lastRow = Worksheets("Лист1").Cells(Rows.Count, 1).End(xlUp).Row</Fix>

        <Error>Использование .Select и .Activate (плохая практика).</Error>
        <BadCode><![CDATA[
Worksheets("Лист1").Range("A1").Select
Selection.Value = "Привет"
        ]]></BadCode>
        <GoodCode><![CDATA[
Worksheets("Лист1").Range("A1").Value = "Привет"
        ]]></GoodCode>

        <Error>Жёстко прописанные пути или имена файлов без возможности выбора.</Error>
        <Fix>Использовать Application.GetOpenFilename или Application.GetSaveAsFilename.</Fix>

        <Error>Ошибки с типами переменных. Нельзя оставлять Dim i (Variant по умолчанию).</Error>
        <Fix>Использовать Dim i As Long.</Fix>
    </CommonErrors>

    <BestPractices>
        <Practice>Разделять код на процедуры и функции: каждая выполняет одну задачу.</Practice>
        <Practice>Использовать конструкцию With ... End With для сокращения повторов.</Practice>
        <Example><![CDATA[
With Worksheets("Лист1")
    .Range("A1").Value = "Тест"
    .Range("B1").Value = "Ок"
End With
        ]]></Example>
        <Practice>Добавлять проверки перед действиями (существует ли лист, не пуст ли диапазон).</Practice>
        <Practice>Добавлять обработку ошибок (On Error GoTo Handler) хотя бы базовую.</Practice>
        <Practice>Делать код гибким: использовать переменные и поиск по имени вместо жёстких ссылок.</Practice>
    </BestPractices>

    <ModelRequirements>
        <Requirement>Код должен запускаться сразу, без недописанных кусков или заглушек.</Requirement>
        <Requirement>Отдавать полный модуль или процедуру, а не разрозненные строки.</Requirement>
        <Requirement>Если есть несколько решений — выбрать наиболее устойчивое и оптимальное, альтернативы давать после.</Requirement>
        <Requirement>В пояснениях избегать общих фраз типа "можно сделать так" без примеров кода.</Requirement>
    </ModelRequirements>
</VBA_Rules>
