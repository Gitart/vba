Option Compare Database
Option Explicit
'*******************************************************************************************
' © Крук Валерий Николаевич, 2007
'*******************************************************************************************
'Сводная таблица + диаграмма в MSExcel
'
Sub StajInExcel()
On Error GoTo Err_StajInExcel
    Dim xlApp As Excel.Application
    Dim xlBook As Excel.Workbook
    Dim xlSheet As Excel.Worksheet
    Dim objPivotCache As Excel.PivotCache
    Dim MyRange As Excel.Range
    Dim rs As New ADODB.Recordset

    rs.CursorLocation = adUseClient 'Рекордсет будет создан у клиента
    rs.Open "SELECT KatName AS Работники, StajPeriod AS [Стаж работы], Count_KATDOLJ AS Количество, DatePokaz AS Год FROM tab_3_Staj", _
        CurrentProject.Connection, adOpenStatic, adLockReadOnly, adCmdText

    Set xlApp = CreateObject("Excel.Application") 'Создание объекта MSExcel
    Set xlBook = xlApp.Workbooks.Add 'Создание файла Excel
    'xlApp.Visible = True 'Выводим на экран (оставлено для возможной отладки)
    xlApp.DisplayAlerts = False 'Запрет возможных сообщений MSExcel

    Set xlSheet = xlBook.Sheets(1)
    With xlSheet
        .Name = "Сводная" 'Присваивем листу имя
        'Создаем сводную таблицу с внешним источником данных (xlExternal)
        Set objPivotCache = xlBook.PivotCaches.Add(xlExternal)
        'Присваиваем сводной таблице в качестве источника данных рекордсет (rs)
        Set objPivotCache.Recordset = rs
        rs.Close 'Закрываем рекордсет, т.к. он больше не нужен
        Set rs = Nothing 'Чистим память от объекта

        'Создаем каркас для сводной и указываем что будет строками, а что столбцами
        .PivotTables.Add PivotCache:=objPivotCache, TableDestination:=.Cells(2, 1), TableName:="Svodnaya"
        With .PivotTables("Svodnaya").PivotFields("Работники")
            .Orientation = xlRowField 'Строка
            .Position = 1 'Позиция №1
        End With
        With .PivotTables("Svodnaya").PivotFields("Стаж работы")
            .Orientation = xlRowField 'Строка
            .Position = 2 'Позиция №2
        End With
        With .PivotTables("Svodnaya").PivotFields("Год")
            .Orientation = xlColumnField 'Столбец
            .Position = 1 'Позиция №1
        End With

        'Подбиваем суммы по группам
        .PivotTables("Svodnaya").AddDataField .PivotTables("Svodnaya").PivotFields("Количество"), "Кол-во", xlSum

        'Тоже, только в процентах (xlPercentOfColumn)
        .PivotTables("Svodnaya").AddDataField .PivotTables("Svodnaya").PivotFields("Количество"), "%", xlSum
        With .PivotTables("Svodnaya").DataPivotField
            .Orientation = xlColumnField 'Говорим сводной, что итоги по процентам нужно вывести в столбец
            .Position = 2 'На позицию №2
        End With
        .PivotTables("Svodnaya").PivotFields("%").Calculation = xlPercentOfColumn

        '=================================================================================
        'Сводная таблица создана!
        '=================================================================================

        'Далее косметика (подписываем таблицу и наводим красивости)
        '1. Подпись сводной
        Set MyRange = .Range(.Cells(1, 1), .Cells(1, 5)) 'Диапазон ячеек для подписи
         MyRange.Merge                                   'Объединяем ячейки
         MyRange.Font.Bold = True                        'Назначаем жирный шрифт
         MyRange = "Стаж работы"                         'Собственно сама подпись
        .Rows(1).RowHeight = 35                          'Устанавливаем высоту строки
        .Rows(1).HorizontalAlignment = xlCenter          'Выравнивание подписи по вертикали
        .Rows(1).VerticalAlignment = xlCenter            'Выравнивание подписи по горизонтали

        '2. Шапка таблицы
        .Rows("3:4").HorizontalAlignment = xlCenter
        .Rows("3:4").VerticalAlignment = xlCenter
        .Rows("3:4").Font.Bold = True
        .Range(.Cells(3, 1), .Cells(4, 2)).Interior.ColorIndex = 40
        .Range(.Cells(2, 3), .Cells(2, 4)).Interior.ColorIndex = xlNone
        .Columns(1).ColumnWidth = 18                                             'Ширина колонки
        
        'Выкрашиваем шапку таблицы в светло-коричневый цвет
        .PivotTables("Svodnaya").PivotSelect "Год[All]", xlLabelOnly, True
         xlApp.Selection.Interior.ColorIndex = 40
        .PivotTables("Svodnaya").PivotSelect "Данные[All]", xlLabelOnly, True
         xlApp.Selection.Interior.ColorIndex = 40
        .Cells(2, 1).Select

        '3. Для того, чтобы визуально отделить столбцы (количество и проценты)
        'красим шрифт подписей и значений процентов в коричневый цвет
        .PivotTables("Svodnaya").PivotFields("%").LabelRange.Font.ColorIndex = 9
        .PivotTables("Svodnaya").PivotFields("%").DataRange.Font.ColorIndex = 9

        '4. Красим ячейки подписей групп
        .PivotTables("Svodnaya").PivotSelect "Работники[All]", xlLabelOnly, True
        xlApp.Selection.Interior.ColorIndex = 35
        xlApp.Selection.Font.Bold = True 'Жирный шрифт для подписей групп

        '5. Подписи и суммы групп
        .PivotTables("Svodnaya").PivotSelect "Работники[All;Total]", xlDataAndLabel, True
        xlApp.Selection.Font.Bold = True 'Выделяем жирным
        .PivotTables("Svodnaya").PivotFields("Работники").SubtotalName = " " 'а сами подписи убираем

        '6. Общие итоги
        .PivotTables("Svodnaya").PivotSelect "'Column Grand Total'", xlDataAndLabel, True
        xlApp.Selection.Font.Bold = True 'Жирным
        xlApp.Selection.Interior.ColorIndex = 40 'Цвет ячеек - светло-коричневый
        .PivotTables("Svodnaya").GrandTotalName = "Общее количество работников" 'Подпись

        '7. Определяем некоторые общие установки сводной таблицы
        With .PivotTables("Svodnaya")
            .HasAutoFormat = True 'Авто формат
            .NullString = "0" 'Вместо NULL значений выводим 0
            .RowGrand = False 'Скрываем колонки с итогами по строкам
        End With

        '=================================================================================
        'Рисуем диаграмму
        '=================================================================================
        'Добавляем диаграмму (тип - xlColumnClustered) на новый лист
        xlApp.Charts.Add
        xlApp.ActiveChart.ChartType = xlColumnClustered
        xlApp.ActiveChart.PlotArea.Interior.ColorIndex = xlNone         'Обесцвечиваем подложку (фон)
        xlApp.ActiveChart.HasTitle = True                               'Отображение заголовка диаграммы
        xlApp.ActiveChart.ChartTitle.Characters.Text = "Стаж работы"    'Заголовок
        xlApp.ActiveChart.Legend.Position = xlTop                       'Вывод легенды сверху диаграммы
        xlApp.ActiveSheet.Name = "Диаграмма"                            'Наименование листа
        'Перемещаем лист с диаграммой на вторую позицию (после листа со сводной таблицей)
        xlBook.Sheets("Диаграмма").Move After:=xlBook.Sheets(2)

        .Select 'Переходим на первый лист
        .Cells(4, 1).Select
    End With

   'Скрываем 'повылазившие' панели инструментов
    xlApp.ActiveWorkbook.ShowPivotTableFieldList = False
    xlApp.CommandBars("PivotTable").Visible = False
    xlApp.CommandBars("Chart").Visible = False

    'Сохранение файла под именем Staj.xls
    xlBook.SaveAs FileName:=CurrentProject.Path & "\Staj", FileFormat:=xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False, CreateBackup:=False
    xlApp.DisplayAlerts = True                  'Разрешаем сообщения MSExcel
    xlApp.Visible = True                        'Выводим на экран

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

Exit Sub
Err_StajInExcel:
    MsgBox Err.Description, vbCritical + vbMsgBoxHelpButton, _
        "Ошибка №" & Err.Number, Err.HelpFile, Err.HelpContext
    On Error Resume Next
    xlApp.Quit
End Sub
