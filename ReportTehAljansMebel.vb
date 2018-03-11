Dim Executor
Dim Report
Dim FlagReport
Dim mdf
Dim BookReport
Dim SignServer
Dim objExcel


Sub Отчет_технолога()
'
' Отчет_технолога Макрос
' Макрос записан 29.02.2016 (AL)
'
    Executor = "Бахирев А.А."
    Report = "D:\работа\автоформатирование\КнигаПрофиля.xls"
    SignServer = "D:\"
    FlagReport = 0
    mdf = 0
'
CopyImage 'копирую картинки

Set objExcel = New Excel.Application
Set BookReport = objExcel.Workbooks.Open(Report)

For i = 4 To 9999
    If BookReport.Sheets(Executor).Cells(i, 2) = ActiveWorkbook.Name Then
        WriteReport (i)
        Exit For
        End If
    If IsEmpty(BookReport.Sheets(Executor).Cells(i, 2)) Then
        WriteReport (i)
        Exit For
        End If
    Next i

    Application.DisplayAlerts = False
    On Error Resume Next
    BookReport.Save
    BookReport.Close ' обязательно при выходе из кода
    Set objExcel = Nothing  ' обязательно при выходе из кода
    
    On Error GoTo 0
    Application.DisplayAlerts = True

End Sub

' процедура присваивания строки

Sub WriteReport(Number)
        
        BookReport.Sheets(Executor).Cells(Number, 1).NumberFormat = "dd.mm.yyyy"
        BookReport.Sheets(Executor).Cells(Number, 1) = Date
        BookReport.Sheets(Executor).Cells(Number, 2) = ActiveWorkbook.Name
        BookReport.Sheets(Executor).Cells(Number, 3) = ActiveWorkbook.Sheets("Калькуляция").Cells(3, 8)
        
        For j = 1 To 9999
            If ActiveWorkbook.Sheets("Калькуляция").Cells(j, 2) = "Себестоимость" Then
                BookReport.Sheets(Executor).Cells(Number, 4) = ActiveWorkbook.Sheets("Калькуляция").Cells(j, 7)
                BookReport.Sheets(Executor).Cells(Number, 4).NumberFormat = "0"
                Exit For
                End If
            Next j
        
        For j = 1 To 9999
            If ActiveWorkbook.Sheets("Калькуляция").Cells(j, 2) = "Закупочная" Then
                BookReport.Sheets(Executor).Cells(Number, 5) = ActiveWorkbook.Sheets("Калькуляция").Cells(j, 7)
                BookReport.Sheets(Executor).Cells(Number, 5).NumberFormat = "0"
                Exit For
                End If
            Next j
        BookReport.Sheets(Executor).Cells(Number, 6) = "уп."
        BookReport.Sheets(Executor).Cells(Number, 7) = Executor
        For j = 1 To 9999
            If ActiveWorkbook.Sheets("Калькуляция").Cells(j, 2) = "Себестоимость" Then
                For l = j To 9999
                    If Left(ActiveWorkbook.Sheets("Калькуляция").Cells(l, 2), 4) = "МДФ " Then
                        mdf = mdf + ActiveWorkbook.Sheets("Калькуляция").Cells(l, 7).Value
                        End If
                    If IsEmpty(ActiveWorkbook.Sheets("Калькуляция").Cells(l, 2)) Then
                        Exit For
                        End If
                    Next l
                Exit For
                End If
            Next j
        If mdf > 0 Then
            BookReport.Sheets(Executor).Cells(Number, 8) = "МДФ " & mdf
            BookReport.Sheets(Executor).Cells(Number, 8).NumberFormat = "0"
            End If
End Sub

Sub CopyImage()
    Application.DisplayAlerts = False
    On Error Resume Next
     
    FileCopy Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, ".") - 1) & ".jpg", SignServer & Left(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & ".jpg"
    FileCopy ActiveWorkbook.Path & "\" & Left(ActiveWorkbook.Name, 5) & "картинка.jpg", SignServer & Left(ActiveWorkbook.Name, 5) & "картинка.jpg"
    
    On Error GoTo 0
    Application.DisplayAlerts = True

End Sub
