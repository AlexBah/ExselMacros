Dim Executor
Dim FlagFur
Dim FlagEvrovint
Dim i
Dim FirstFur
Dim j
Dim Temp
Dim FlagMDF
Dim FlagGlaz
Dim FlagCloth
Dim FlagSpanbond
Dim Nzakup
Dim MDFcount

   
   

Sub автоформатирование()
' автоформатирование Макрос
' Макрос записан 11.02.2016 (AL)

' определяю константы
    Executor = "Бахирев А.А."

' Форматирую листы
    Sheets("Листовые детали").Select
    If Cells(1, 3) = "Код" Then
        Columns("C:C").Cut Destination:=Columns("J:J")
        Columns("C:C").Delete Shift:=xlToLeft
        End If
    Columns("H:H").Cut Destination:=Columns("E:E")
    Columns("H:H").Delete
    Sheets("Расход материала").Select
    Columns("D:E").Delete
    Columns("I:K").Delete
    FlagFur = 0
    FlagEvrovint = 0

' цикл для уничтожения не нужных строк и подсчета количества деталей и добавления ключика
    For i = 1 To 1000
        If Cells(i, 1) = "Крепёж" Then
            FirstFur = i: FlagFur = 1
            End If
        If Cells(i, 2) = "Евровинт 7*50" Then
            FlagEvrovint = 1
            End If
        If Left(Cells(i, 2), 5) = "Ткань" Then
            Cells(i, 3).Value = "метр пог"
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 1.4, 0)
            Cells(i, 4).NumberFormat = "0.00"
            Cells(i, 5).NumberFormat = "0.00"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 1.4, 2) ' 205
            End If
        If Left(Cells(i, 2), 22) = "Спанбонд" Then
            Cells(i, 3).Value = "метр пог"
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 1.5, 0)
            Cells(i, 4).NumberFormat = "0.00"
            Cells(i, 5).NumberFormat = "0.00"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 1.5, 2) ' 205
            End If
        If Left(Cells(i, 2), 22) = "Профильный погонаж №17" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2.4, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).NumberFormat = "0.00"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2.4, 2) ' 205
            End If
        If Left(Cells(i, 2), 12) = "Планка №31/4" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2.4, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).NumberFormat = "0.00"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2.4, 2) ' 166.5
            End If
        If Left(Cells(i, 2), 10) = "Планка №38" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2.4, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2.4, 2) ' 90
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль торц.ЛДСП 16мм круг" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 186.43
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль торц.ЛДСП 16мм" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 186.43
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Соединитель для ДВПО" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2, 2) ' 22.1
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Шина 2 м." Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2, 2) ' 22.1
            Cells(i, 5).NumberFormat = "0.00"
            End If
'        If Cells(i, 2) = "НАЙДИ Напр-я верх Хром мат" Or Cells(i, 2) = "НАЙДИ Напр-я низ Хром мат" Or Cells(i, 2) = "НАЙДИ Профиль верх Хром мат" Or Cells(i, 2) = "НАЙДИ Профиль низ Хром мат" Or Cells(i, 2) = "НАЙДИ Профиль средн Хром мат" Then
'            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4), 1)
'            If Cells(i, 4).Value < 1.2 Then
'                Cells(i, 4).Value = 1.2
'                End If
'            Cells(i, 4).NumberFormat = "0.0"
'            End If
        If Left(Cells(i, 2), 19) = "НАЙДИ Профиль гориз" Or Left(Cells(i, 2), 12) = "НАЙДИ Напр-я" Then
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4), 1)
            If Cells(i, 4).Value < 1.2 Then
                Cells(i, 4).Value = 1.2
                End If
            Cells(i, 4).NumberFormat = "0.0"
            End If
         If Cells(i, 2) = "КЛАССИК Напр-я верх Хром мат" Or Cells(i, 2) = "КЛАССИК Напр-я низ Хром мат" Or Cells(i, 2) = "КЛАССИК Профиль верх Хром мат" Or Cells(i, 2) = "КЛАССИК Профиль низ Хром мат" Or Cells(i, 2) = "КЛАССИК Профиль средн Хром мат" Then
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4), 1)
            If Cells(i, 4).Value < 1.2 Then
                Cells(i, 4).Value = 1.2
                End If
            Cells(i, 4).NumberFormat = "0.0"
            End If
        
        If Cells(i, 2) = "Профиль RAUM гориз. нижний 3м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 869
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль RAUM гориз. верхний 3м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 432
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль RAUM гориз. нижний 2м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2, 2) ' 579
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль RAUM гориз. средний 3м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 406
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль RAUM гориз. верхний 2м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2, 2) ' 288
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Cells(i, 2) = "Профиль RAUM гориз. средний 2м" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 2, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 2, 2) ' 271
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Left(Cells(i, 2), 13) = "Щит мебельный" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 1257.45
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Left(Cells(i, 2), 16) = "Столешница 38 мм" Then
            Cells(i, 3).Value = "шт."
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3, 0)
            Cells(i, 4).NumberFormat = "0"
            Cells(i, 5).Value = Application.RoundUp(Cells(i, 5).Value * 3, 2) ' 2008.8
            Cells(i, 5).NumberFormat = "0.00"
            End If
        If Left(Cells(i, 2), 11) = "Кромка 3,05" Then
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3.05, 0) * 3.05
            Cells(i, 4).NumberFormat = "0.00"
            End If
        If Left(Cells(i, 2), 18) = "Кромка 3,05 Влаг-я" Then
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4) / 3.05, 0) * 3.05
            Cells(i, 4).NumberFormat = "0.00"
            End If

        If IsEmpty(Worksheets(ActiveSheet.Name).Cells(i, 1)) = True Then
            DetailsCount = i - 1
            If FlagEvrovint = 1 Then
                Cells(i, 1).Value = i
                Cells(i, 2).Value = "Ключ д/евровинта 4 мм"
                Cells(i, 3).Value = "шт."
                Cells(i, 4).Value = "1"
                Cells(i, 5).Value = 1.44
                Cells(i, 6).Value = "руб."
                Cells(i, 7).Value = 1.44
                Cells(i, 8).Value = "руб."
                DetailsCount = DetailsCount + 1
                End If
            Exit For
            End If
        If Cells(i, 1) = "№" Or IsEmpty(Cells(i, 2)) = True Or Cells(i, 2) = "Клей расплав" Or Cells(i, 2) = "Ключ д/евровинта 4 мм" Then
            Rows(i).Delete: i = i - 1
            Else
            'округляю до второго знака после запятой
            Cells(i, 4).Value = Application.RoundUp(Cells(i, 4).Value, 2)
            End If
        Next i
' сортировка по названию и новая нумерация
   If FlagFur = 1 Then
        Range(Cells(FirstFur, 1), Cells(DetailsCount, 8)).Select
        Selection.Sort Key1:=Range(Cells(FirstFur, 2), Cells(FirstFur, 2)), Order1:=xlAscending, Header:=xlGuess, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal
        End If
    For i = 1 To DetailsCount
        Cells(i, 1) = i
        Next i

' формирование шапки
Rows("1:4").Select
    Selection.Insert Shift:=xlDown
Range("A1:H1").Select
    MergeCenter
    Cells(1, 1) = "ООО Альянс Мебель"
Range("A2:H2").Select
    MergeCenter
    Cells(2, 1) = "Изделие: заказ №" & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1)
Range("A3:g3").Select
    MergeCenter
    Cells(3, 1) = "Калькуляция"
    Cells(3, 8) = 1.4
    
    Cells(4, 1) = "№"
    Cells(4, 2) = "Наименование"
    Cells(4, 3) = "Ед.изм."
    Cells(4, 4) = "Кол."
    Cells(4, 5) = "Цена за ед."
    Cells(4, 7) = "Цена в исх.валюте"
    Range("A1:H4").Font.Bold = True
        
'создание границ
    Range(Cells(1, 1), Cells(DetailsCount + 14, 8)).Select
    ThinBorders
    For i = 1 To 5
        If i = 1 Then
            Range("A1:H3").Select
            End If
        If i = 2 Then
            Range("A4:H4").Select
            End If
        If i = 3 Then
            Range(Cells(5, 1), Cells(DetailsCount + 5, 8)).Select
            End If
        If i = 4 Then
            Range(Cells(DetailsCount + 6, 1), Cells(DetailsCount + 10, 8)).Select
            End If
        If i = 5 Then
            Range(Cells(DetailsCount + 11, 1), Cells(DetailsCount + 14, 8)).Select
            End If
        ThickBorders
        Next i
    Range("A1:H3").Interior.ColorIndex = 6
    
    For i = 1 To DetailsCount
        Cells(i + 4, 7).FormulaR1C1 = "=RC[-3]*RC[-2]"
        Cells(i, 7).NumberFormat = "0.00"
        Next i
        
'заполнение постскрипта
        Cells(DetailsCount + 6, 2).Value = "Итого"
        Temp = "=SUM(R[-" & 1 + DetailsCount & "]C7:R[-1]C7)"
        Cells(DetailsCount + 6, 7).FormulaR1C1 = Temp
        Cells(DetailsCount + 7, 2).Value = "Работа"
        Cells(DetailsCount + 7, 7).FormulaR1C1 = "=Roundup(R[-1]C7*0.16 , -1)"
        Cells(DetailsCount + 8, 2).Value = "Упаковка"
        Temp = "=0"
        For i = 1 To DetailsCount
            If Left(Cells(i + 4, 2).Value, 4) = "ЛДСП" Then
                Temp = Temp & "+R" & i + 4 & "C"
                End If
            Next i
        Cells(DetailsCount + 8, 4).FormulaR1C1 = Temp
        Cells(DetailsCount + 8, 5).Value = 15
        Cells(DetailsCount + 8, 7).FormulaR1C1 = "=RC[-3]*RC[-2]"
        Cells(DetailsCount + 9, 2).Value = "Себестоимость"
        Cells(DetailsCount + 9, 7).FormulaR1C1 = "=SUM(R[-3]C7:R[-1]C7)"
        Cells(DetailsCount + 11, 2).Value = "Закупочная"
        Cells(DetailsCount + 11, 7).FormulaR1C1 = "=R[-2]C*R3C8"
        Cells(DetailsCount + 12, 2).Value = "Оптовая"
        Cells(DetailsCount + 12, 7).FormulaR1C1 = "=R[-1]C*1.1"
        Cells(DetailsCount + 13, 2).Value = "Мелкооптовая"
        Cells(DetailsCount + 13, 7).FormulaR1C1 = "=R[-2]C*1.17"
        Cells(DetailsCount + 14, 2).Value = "Розничная"
        Cells(DetailsCount + 14, 7).FormulaR1C1 = "=R[-3]C*1.35"
        Cells(DetailsCount + 15, 2).Value = "Конструктор: " & Executor
        Cells(DetailsCount + 16, 2).Value = "Монтаж"
        Cells(DetailsCount + 17, 2).Value = "Такси"
        Range(Cells(DetailsCount + 16, 1), Cells(DetailsCount + 17, 8)).Select
        ThickBorders
        
    For i = 5 To DetailsCount + 14
        Cells(i, 8).Value = "руб."
        Next i
    For i = 1 To 9
        Cells(DetailsCount + 5 + i, 7).NumberFormat = "0"
        Next i


'выравнивание по ширине столбцов, разметка страницы
ForPrint





' ************************************************
' Вторая часть "Спецификация"
    
    Sheets("Листовые детали").Select
    
' Подсчет листовых деталей, сортировка и нумерация
    For i = 1 To 1000
        If IsEmpty(Worksheets(ActiveSheet.Name).Cells(i, 1)) = True Then
            listscount = i - 2
            Exit For
            End If
        Next i
    
    Range("A2:H" & listscount + 1).Select
    Selection.Sort Key1:=Range("G2"), Order1:=xlAscending, Key2:=Range("H2"), _
        Order2:=xlAscending, Header:=xlNo, _
        OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        DataOption1:=xlSortNormal

    For i = 1 To listscount
        Cells(i + 1, 1) = i
        Next i

'создание шапки

Rows("1:3").Select
    Selection.Insert Shift:=xlDown
    
Range("A1:H1").Select
    MergeCenter
    Cells(1, 1) = "ООО Альянс Мебель"
Range("A2:H2").Select
    MergeCenter
    Cells(2, 1).FormulaR1C1 = "='Расход материала'!R2C1"
Range("A3:g3").Select
    MergeCenter
    Cells(3, 1) = "Спецификация"
    Cells(3, 8).FormulaR1C1 = "='Расход материала'!R3C8"
    Cells(3, 8).NumberFormat = "0.0"
    
    Range("A1:H4").Select
    Selection.Font.Bold = True

Range("A" & listscount + 5 & ":h" & listscount + 5).Select
    MergeCenter
    Cells(listscount + 5, 1) = "Фурнитура"
    Selection.Font.Bold = True
    
' создание тонких границ
    Range(Cells(1, 1), Cells(DetailsCount + listscount + 7, 8)).Select
    ThinBorders
' создание толстых границ
    For i = 1 To 7
        If i = 1 Then
            Range("A1:H3").Select
            End If
        If i = 2 Then
            Range("A4:H4").Select
            End If
        If i = 3 Then
            Range(Cells(5, 1), Cells(listscount + 4, 8)).Select
            End If
        If i = 4 Then
            Range(Cells(listscount + 5, 1), Cells(listscount + 5, 8)).Select
            End If
        If i = 5 Then
            Range(Cells(listscount + 6, 1), Cells(listscount + 6, 8)).Select
            End If
        If i = 6 Then
            Range(Cells(listscount + 7, 1), Cells(listscount + DetailsCount + 6, 8)).Select
            End If
        If i = 7 Then
            Range(Cells(listscount + DetailsCount + 7, 1), Cells(listscount + DetailsCount + 7, 8)).Select
            End If
        ThickBorders
        Next i
    
Range("A1:H3").Interior.ColorIndex = 6
Range(Cells(listscount + 5, 1), Cells(listscount + 5, 8)).Interior.ColorIndex = 6

Temp = 0
For i = 1 To listscount
    If Left(Cells(i + 4, 7), 7) = "Стекло " Or Left(Cells(i + 4, 7), 8) = "Зеркало " Then
        Range(Cells(i + 4, 1), Cells(i + 4, 8)).Interior.ColorIndex = 6
        End If
    If Cells(i + 4, 7) <> Temp Then
        Range(Cells(i + 4, 1), Cells(i + 4, 8)).Borders(xlEdgeTop).Weight = xlMedium
        Temp = Cells(i + 4, 7)
        End If
    Next i


'перенос из калькуляции
Temp = 0
For i = 1 To DetailsCount + 1
    If Left(Sheets("Расход материала").Cells(i + 3, 2), 9) <> "Отверстие" Then
        For j = 1 To 4
            Cells(listscount + 5 + i - Temp, j).FormulaR1C1 = "='Расход материала'!R" & i + 3 & "C" & j
            Next j
        Else
        Temp = Temp + 1
        End If
    Next i
Cells(listscount + DetailsCount + 7 - Temp, 2).Value = "Работа"
Cells(listscount + DetailsCount + 7 - Temp, 4).FormulaR1C1 = "='Расход материала'!R" & DetailsCount + 7 & "C7"
Cells(listscount + DetailsCount + 8 - Temp, 2).FormulaR1C1 = "='Расход материала'!R" & DetailsCount + 15 & "C2"
Cells(listscount + DetailsCount + 9 - Temp, 2).FormulaR1C1 = "='Расход материала'!R" & DetailsCount + 16 & "C2"
Cells(listscount + DetailsCount + 9 - Temp, 4).FormulaR1C1 = "='Расход материала'!R" & DetailsCount + 16 & "C7"
Cells(listscount + DetailsCount + 10 - Temp, 2).Value = "Такси"
Cells(listscount + DetailsCount + 10 - Temp, 4).FormulaR1C1 = "='Расход материала'!R" & DetailsCount + 17 & "C7"
Range(Cells(listscount + DetailsCount + 9 - Temp, 1), Cells(listscount + DetailsCount + 10 - Temp, 8)).Select
ThickBorders

' выравнивание по ширине столбцов, разметка страницы, переименование листов
ForPrint
'вставляю рисунок
Application.DisplayAlerts = False
On Error Resume Next
Cells(listscount + DetailsCount + 12 - Temp, 1).Select
ActiveSheet.Pictures.Insert(Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, ".") - 1) & ".jpg").Select
On Error GoTo 0
Application.DisplayAlerts = True
    
    
' переименование листов и удаление
    Sheets("Расход материала").Select
    Sheets("Расход материала").Name = "Калькуляция"
    Sheets("Листовые детали").Name = "Спецификация"
    Application.DisplayAlerts = False
    On Error Resume Next
    Sheets("Кромки").Delete
    Sheets("Крепёжные детали").Delete
    Sheets("Комплектующие").Delete
    Sheets("Профили").Delete
    Sheets("Материалы профилей").Delete
    Sheets("Лист3").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
' работа с МДФ
    FlagMDF = 0
    FlagGlaz = 0
    FlagCloth = 0
    FlagSpanbond = 0
    Sheets("Калькуляция").Select
    For i = 1 To 1000
        If Cells(i, 2) = "Закупочная" Then
            Nzakup = i
            Exit For
            End If
        Next i
    
        
    For i = 1 To DetailsCount
        If Left(Cells(i + 4, 2), 7) = "Стекло " Or Left(Cells(i + 4, 2), 8) = "Зеркало " Then
            FlagGlaz = 1
            Rows(Nzakup).Insert Shift:=xlDown
            Cells(Nzakup - 1, 2).FormulaR1C1 = "=R" & i + 4 & "C2"
            Cells(Nzakup - 1, 2).Font.Bold = True
            Nzakup = Nzakup + 1
            Temp = Cells(Nzakup - 2, 2)
            For j = 1 To listscount
                If Sheets("Спецификация").Cells(j + 4, 7) = Temp Then
                    Rows(Nzakup).Insert Shift:=xlDown
                    Nzakup = Nzakup + 1
                    Cells(Nzakup - 2, 2).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C2"
                    Cells(Nzakup - 2, 3).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C3"
                    Cells(Nzakup - 2, 4).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C4"
                    Cells(Nzakup - 2, 5).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C5"
                    End If
                Next j
            End If
        Next i
    
    For i = 1 To DetailsCount
        If Left(Cells(i + 4, 2), 5) = "Ткань" Then
            Rows(Nzakup).Insert Shift:=xlDown
            Rows(i + 4).Cut Destination:=Rows(Nzakup - 1)
            Rows(i + 4).Delete
            i = i - 1
            Cells(Nzakup, 7) = Cells(Nzakup, 7).FormulaR1C1 & "+R" & Nzakup - 2 & "C"
            Cells(Nzakup - 2, 2).Font.Bold = True
            If FlagCloth = 0 Then
                FlagCloth = 1
                For j = 1 To listscount
                    If Left(Sheets("Спецификация").Cells(j + 4, 7), 5) = "Ткань" Then
                        Sheets("Спецификация").Cells(j + 4, 3) = Sheets("Спецификация").Cells(j + 4, 3) + 118
                        Sheets("Спецификация").Cells(j + 4, 4) = Sheets("Спецификация").Cells(j + 4, 4) + 118
                        End If
                    Next j
                End If
            End If
        
        If (Left(Cells(i + 4, 2), 8) = "Спанбонд") And (FlagSpanbond = 0) Then
            FlagSpanbond = 1
            For j = 1 To listscount
                If Left(Sheets("Спецификация").Cells(j + 4, 7), 8) = "Спанбонд" Then
                    Sheets("Спецификация").Cells(j + 4, 3) = Sheets("Спецификация").Cells(j + 4, 3) + 20
                    Sheets("Спецификация").Cells(j + 4, 4) = Sheets("Спецификация").Cells(j + 4, 4) + 20
                    End If
                Next j
            End If
        Next i
    
    MDFcount = 6
    For i = 1 To DetailsCount
        If Left(Cells(i + 4, 2), 3) = "МДФ" Then
            If FlagMDF = 0 Then
                Worksheets.Add.Name = "МДФ"
                Sheets("МДФ").Range("A1:G1").Select
                MergeCenter
                Cells(1, 1) = "Для цеха Альянс Мебели от " & CStr(Date)
                Sheets("МДФ").Range("A2:G2").Select
                MergeCenter
                Cells(2, 1).FormulaR1C1 = "='Калькуляция'!R2C1"
                Cells(3, 2) = "Название"
                Cells(3, 3) = "Цвет"
                Cells(3, 4) = "Рисунок"
                Cells(3, 5) = "Высота"
                Cells(3, 6) = "Ширина"
                Cells(3, 7) = "Кол."
                Sheets("МДФ").Range("A1:G3").Font.Bold = True
                FlagMDF = 4
                Sheets("Калькуляция").Select
                End If
            Rows(Nzakup).Insert Shift:=xlDown
            Rows(i + 4).Cut Destination:=Rows(Nzakup - 1)
            Rows(i + 4).Delete
            i = i - 1
            Cells(Nzakup, 7) = Cells(Nzakup, 7).FormulaR1C1 & "+R" & Nzakup - 2 & "C"
            Cells(Nzakup - 2, 2).Font.Bold = True
            Temp = Cells(Nzakup - 2, 2)
            ' забиваю пленку МДФ и указываю последнюю пустую строчку
            Sheets("МДФ").Cells(MDFcount, 1) = "пленка УралПлит, м."
            Sheets("МДФ").Cells(MDFcount, 3) = Temp
            MDFcount = MDFcount + 1
            For j = 1 To listscount
                If Sheets("Спецификация").Cells(j + 4, 7) = Temp Then
                    Rows(Nzakup).Insert Shift:=xlDown
                    Nzakup = Nzakup + 1
                    Cells(Nzakup - 2, 2).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C2"
                    Cells(Nzakup - 2, 3).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C3"
                    Cells(Nzakup - 2, 4).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C4"
                    Cells(Nzakup - 2, 5).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C5"
                    Sheets("МДФ").Rows(MDFcount - 2).Insert Shift:=xlDown
                    Sheets("МДФ").Cells(FlagMDF, 2).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C2"
                    Sheets("МДФ").Cells(FlagMDF, 3).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C7"
                    Sheets("МДФ").Cells(FlagMDF, 4).FormulaR1C1 = "=RIGHT(R" & FlagMDF & "C2,4)"
                    Sheets("МДФ").Cells(FlagMDF, 5).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C3"
                    Sheets("МДФ").Cells(FlagMDF, 6).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C4"
                    Sheets("МДФ").Cells(FlagMDF, 7).FormulaR1C1 = "='Спецификация'!R" & j + 4 & "C5"
                    FlagMDF = FlagMDF + 1
                    MDFcount = MDFcount + 1
                    End If
                Next j
            End If
        Next i

'вставляю работу швеи и обтяжку
If FlagCloth > 0 Then
    Sheets("Калькуляция").Select
    For i = 1 To 1000
        If Cells(i, 2) = "Работа" Then
            Rows(i + 1).Insert Shift:=xlDown
            Cells(i + 1, 2) = "Работа швея"
            Cells(i + 1, 8) = "руб."
            Rows(i + 1).Insert Shift:=xlDown
            Cells(i + 1, 2) = "Работа обтяжка"
            Cells(i + 1, 8) = "руб."
            Temp = i
            Exit For
            End If
        Next i
    Sheets("Спецификация").Select
    For i = 1 To 1000
        If Cells(i, 2) = "Работа" Then
            Rows(i + 1).Insert Shift:=xlDown
            Cells(i + 1, 2) = "Работа швея"
            Cells(i + 1, 4).FormulaR1C1 = "='Калькуляция'!R" & Temp + 2 & "C7"
            Rows(i + 1).Insert Shift:=xlDown
            Cells(i + 1, 2) = "Работа обтяжка"
            Cells(i + 1, 4).FormulaR1C1 = "='Калькуляция'!R" & Temp + 1 & "C7"
            Range(Cells(i, 1), Cells(i + 2, 8)).Select
            ThinBorders
            ThickBorders
            Exit For
            End If
        Next i
        
    End If

' форматирование МДФ и стекла
If FlagMDF > 0 Or FlagGlaz > 0 Or FlagCloth > 0 Then
    Sheets("Калькуляция").Select
    For i = 1 To 1000
        If Cells(i, 2) = "Себестоимость" Then
            Startformat = i + 1
            End If
        If Cells(i, 2) = "Закупочная" Then
            Endformat = i - 1
            Exit For
            End If
        Next i
    Range(Cells(Startformat, 1), Cells(Endformat, 8)).Select
    ThinBorders
    ThickBorders
    Range(Cells(Startformat, 1), Cells(Endformat, 8)).Interior.ColorIndex = 6
    End If
    
    
'форматирование страницы МДФ
If FlagMDF > 0 Then
' создание тонких границ
    Sheets("МДФ").Select
    Range(Cells(3, 1), Cells(FlagMDF - 1, 7)).Select
    ThinBorders
' выравнивание по ширине столбцов, разметка страницы, переименование листов
    Sheets("МДФ").Columns("A:G").EntireColumn.AutoFit
    Sheets("МДФ").Select
    Sheets("МДФ").Move After:=Sheets(3)
    End If
Sheets("Калькуляция").Select
Sheets("Калькуляция").Move Before:=Sheets(1)

End Sub

Sub ThinBorders()
'создание тонких границ
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Sub ThickBorders()
'создание толстых границ
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlMedium
            .ColorIndex = xlAutomatic
        End With
End Sub

Sub MergeCenter()
'слить и отцентровать
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
End Sub

Sub ForPrint()
'для печати
Cells.Font.Size = 14
Columns("A:H").EntireColumn.AutoFit
With ActiveSheet.PageSetup
    .LeftMargin = Application.InchesToPoints(0.196850393700787)
    .RightMargin = Application.InchesToPoints(0.196850393700787)
    .TopMargin = Application.InchesToPoints(0.196850393700787)
    .BottomMargin = Application.InchesToPoints(0.196850393700787)
    .HeaderMargin = Application.InchesToPoints(0.196850393700787)
    .FooterMargin = Application.InchesToPoints(0.196850393700787)
    End With
ActiveWindow.View = xlPageBreakPreview
Application.DisplayAlerts = False
On Error Resume Next
ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
On Error GoTo 0
Application.DisplayAlerts = True
ActiveWindow.Zoom = 100
End Sub

