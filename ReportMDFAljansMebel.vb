Sub Отчет_МДФ()

FilePath = "D:"
FileN = FilePath & "\" & ActiveWorkbook.Name
Report = "D:\работа\database\Отчет по МДФ.xls"
Executor = "Бахирев А.А."

Application.DisplayAlerts = False
ActiveWorkbook.Sheets("МДФ").Copy
'убираю ссылки
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
'сохраняю файл
ActiveWorkbook.SaveCopyAs FileN
ActiveWorkbook.Close SaveChanges:=False
Application.DisplayAlerts = True

' пишу отчет о том когда скидывал МДФ
Set objExcel = New Excel.Application
Set BookReport = objExcel.Workbooks.Open(Report)

For i = 1 To 9999
    If IsEmpty(BookReport.Sheets("Лист1").Cells(i, 1)) Then
        BookReport.Sheets("Лист1").Cells(i, 1).NumberFormat = "dd.mm.yyyy"
        BookReport.Sheets("Лист1").Cells(i, 1) = Date
        BookReport.Sheets("Лист1").Cells(i, 2) = ActiveWorkbook.Name
        BookReport.Sheets("Лист1").Cells(i, 3) = Executor
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

