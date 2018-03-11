Dim Constructor
Dim Zakazchik
Dim Zakaz
Dim Ukazatel
Dim Furniture
Dim EndFurniture

Sub ФорматированиеСметы()
    ПодготовкаПеременных
    Подготовка
    Ukazatel = 1
    ВставкаШапки
    ВставкаСклада1
    Ukazatel = Furniture
    ВставкаПодписи
    Range(Cells(8, 1), Cells(Ukazatel - 1, 3)).Select
    ТолстаяРамка
    Ukazatel = Furniture
    ВставкаШапки2
    ВставкаСклада2
    Ukazatel = Furniture
    ВставкаПодписи
    Ukazatel = Furniture
    ВставкаШапки2
    ВставкаСклада3
    Ukazatel = EndFurniture + 1
    ВставкаПодписи
    Range(Cells(Furniture - 8, 1), Cells(EndFurniture - 8, 3)).Select
    ТолстаяРамка
    ПодготовкаКПечати
End Sub
Sub ПодготовкаПеременных()
    Constructor = "Бахирев А.А."
    Zakaz = Left(ActiveWorkbook.Name, 6)
    Temp = Right(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 7)
    Zakazchik = Left(Temp, InStr(Temp, "_") - 1)
End Sub
Sub ВставкаШапки()
    Rows(Ukazatel & ":" & Ukazatel + 6).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Furniture = Furniture + 7
    EndFurniture = EndFurniture + 7
    СлитьСтроку
    Cells(Ukazatel, 1) = "Заказчик: " & Zakazchik
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1) = "Код: " & Zakaz
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1) = "Наименование: "
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1) = "Стоимость: "
End Sub
Sub ВставкаШапки2()
    Rows(Ukazatel & ":" & Ukazatel + 6).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Furniture = Furniture + 7
    EndFurniture = EndFurniture + 7
    СлитьСтроку
    Cells(Ukazatel, 1).NumberFormat = "0.00"
    Cells(Ukazatel, 1).FormulaR1C1 = "=R1C"
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1).NumberFormat = "0.00"
    Cells(Ukazatel, 1).FormulaR1C1 = "=R2C"
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1).NumberFormat = "0.00"
    Cells(Ukazatel, 1).FormulaR1C1 = "=R3C"
    Ukazatel = Ukazatel + 1
    СлитьСтроку
    Cells(Ukazatel, 1).NumberFormat = "0.00"
    Cells(Ukazatel, 1).FormulaR1C1 = "=R4C"
End Sub
Sub ВставкаСклада1()
    Ukazatel = Ukazatel + 2
    СлитьСтроку
    Cells(Ukazatel, 1) = "Склад плитных материалов"
    Cells(Ukazatel, 1).Select
    ТолстаяРамка
    Ukazatel = Ukazatel + 1
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Font.Bold = True
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).HorizontalAlignment = xlCenter
    Cells(Ukazatel, 1) = "Ответственный: Николайчук Т.В."
    Cells(Ukazatel, 2) = "Заказано"
    Cells(Ukazatel, 3) = "Получено"
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Select
    ТолстаяРамка
End Sub
Sub ВставкаСклада2()
    Ukazatel = Ukazatel + 2
    СлитьСтроку
    Cells(Ukazatel, 1) = "Склад участка заказной мебели"
    Cells(Ukazatel, 1).Select
    ТолстаяРамка
    Ukazatel = Ukazatel + 1
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Font.Bold = True
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).HorizontalAlignment = xlCenter
    Cells(Ukazatel, 1) = "Ответственный: Николайчук Т.В."
    Cells(Ukazatel, 2) = "Заказано"
    Cells(Ukazatel, 3) = "Получено"
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Select
    ТолстаяРамка
    Ukazatel = Ukazatel + 1
    Rows(Ukazatel & ":" & Ukazatel + 3).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Furniture = Furniture + 4
    EndFurniture = EndFurniture + 4
    Range(Cells(Ukazatel, 1), Cells(Ukazatel + 3, 3)).Select
    ТолстаяРамка
End Sub
Sub ВставкаСклада3()
    Ukazatel = Ukazatel + 2
    СлитьСтроку
    Cells(Ukazatel, 1) = "Склад ТМЦ ЗАО БМФ"
    Cells(Ukazatel, 1).Select
    ТолстаяРамка
    Ukazatel = Ukazatel + 1
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Font.Bold = True
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).HorizontalAlignment = xlCenter
    Cells(Ukazatel, 1) = "Ответственный: Попова Л.П."
    Cells(Ukazatel, 2) = "Заказано"
    Cells(Ukazatel, 3) = "Получено"
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Select
    ТолстаяРамка
End Sub
Sub СлитьСтроку()
    Range(Cells(Ukazatel, 1), Cells(Ukazatel, 3)).Select
    With Selection
        .Font.Bold = True
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
Sub ТолстаяРамка()
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
    Selection.Borders(xlInsideVertical).Weight = xlMedium
    Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
    Selection.Borders(xlInsideHorizontal).Weight = xlThin
End Sub
Sub Подготовка()
    FlagF = 2
    Columns("C:D").Delete
    For i = 1 To 1000
        If Cells(i, 1) = "Крепеж/Евровинт/Евровинт" Then
            Rows(i + 1).Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Cells(i + 1, 1) = "Ключ для евровинта"
            Cells(i + 1, 2) = "1 шт."
            End If
        If Cells(i, 2) = "" And FlagF = 0 Then
            Furniture = i
            FlagF = 1
            End If
        If Cells(i, 1) = "Панели: Материалы основы" Then
            FlagF = 0
            End If
        If Cells(i, 1) = "" Then
            EndFurniture = i
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Rows(i).Delete
            Cells(i, 1) = "Шильд"
            Cells(i, 2) = "1 шт."
            Exit For
            End If
        If Left(Cells(i, 1), 10) = "Пленка ПВХ" And Right(Cells(i, 2), 4) <> " кв." Then
            Rows(i).Delete
            i = i - 1
            End If
        If Cells(i, 2) = "" Then
            Rows(i).Delete
            i = i - 1
            End If
        Next i
End Sub
Sub ВставкаПодписи()
    Rows(Ukazatel & ":" & Ukazatel + 7).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Furniture = Furniture + 8
    EndFurniture = EndFurniture + 8
    Range(Cells(Ukazatel + 1, 1), Cells(Ukazatel + 3, 3)).Select
        With Selection
        .Font.Bold = False
        .HorizontalAlignment = xlLeft
        End With
    Cells(Ukazatel + 1, 1) = "Конструктор: ________________ /" & Constructor & "/"
    Cells(Ukazatel + 1, 2) = "Получил __________________"
    Cells(Ukazatel + 3, 1) = "Отпустил ________________"
    Cells(Ukazatel + 3, 2) = "Получил __________________"
    Range(Cells(Ukazatel + 4, 1), Cells(Ukazatel + 4, 3)).Select
        With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
        End With
End Sub
Sub ПодготовкаКПечати()
    ActiveWindow.View = xlPageBreakPreview
    Columns("A:A").ColumnWidth = 60
    Columns("B:B").ColumnWidth = 13
    Columns("C:C").ColumnWidth = 12
    ActiveWindow.Zoom = 130
End Sub

