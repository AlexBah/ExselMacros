Sub Экспорт_в_Астру()
   Dim RetVal
   RetVal = Shell(Chr(34) & "C:\Program Files (x86)\Astra\Astra.exe" & Chr(34) & " " & Chr(34) & ActiveWorkbook.FullName & Chr(34) & " -i", 1)
End Sub

