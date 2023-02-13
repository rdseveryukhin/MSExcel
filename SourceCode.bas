Attribute VB_Name = "Module1"
Sub loading()

Dim fso, myPath, myFolder, myFile, myFiles()
Dim lastrow As Long ' - actual amount of rows

'Clearing old data
Workbooks("materials_loading.xlsm").Worksheets(1).Cells(1, 10) = ""
lastrow = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Rows.Count, "A").End(xlUp).Row
For Row = 3 To lastrow
    Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Row, 1) = ""
    Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Row, 1).Borders.LineStyle = False
    Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Row, 2) = ""
    Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Row, 2).Borders.LineStyle = False
Next Row

'' Путь к папке с файлами
myPath = Range("B1")

Set fso = CreateObject("Scripting.FileSystemObject")
Set myFolder = fso.GetFolder(myPath)

'' Если нет Менеджеров в папке
If myFolder.Files.Count = 0 Then
    MsgBox "Нет файлов"
    Exit Sub
Else
    Dim count_files As Integer
    count_files = myFolder.Files.Count
    
    '' Массив с именами менеджеров
    ReDim myFiles(1 To count_files)

    '' Заполняем массив знпчениями путей к файлам
    For Each myFile In myFolder.Files
        i = i + 1
        myFiles(i) = myFile.Path
    Next
    
    '' Программа обработки
    For i = 1 To count_files
        Workbooks.Open (myFiles(i))
        Name = Right(myFiles(i), Len(myFiles(i)) - InStrRev(myFiles(i), "\", , 1))
        
        ''Поиск одного из названий раздела
        j = 1
        While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Раздел 2. Материалы и оборудование в текущих ценах" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Раздел№1. Материалы и оборудование" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Раздел №1. Материалы и оборудование"
            j = j + 1
        Wend
        
        ''Андерсен
        If Workbooks(Name).Worksheets(1).Cells(j, 1) = "Раздел 2. Материалы и оборудование в текущих ценах" Then
            j = j + 1
            While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Итоги по акту:" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Итого по разделу 2 Материалы и оборудование в текущих ценах" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "ИТОГИ ПО АКТУ:"
                If Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Нижний ярус" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Секция 5.6" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Секция 5.5" Then
                    lastrow = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Rows.Count, "A").End(xlUp).Row
                    k = 3
                    flag = False
                    While flag = False ' - iteration between rows main pages
                        If Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = Workbooks(Name).Worksheets(1).Cells(j, 4) Then
                            flag = True
                            If Workbooks(Name).Worksheets(1).Cells(j, 6) = "" Then
                                Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) + Workbooks(Name).Worksheets(1).Cells(j, 8)
                            Else
                                Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) + Workbooks(Name).Worksheets(1).Cells(j, 6)
                            End If
        
                        ElseIf Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = "" Or k > lastrow Then
                            flag = True
                            k = lastrow
                            Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).Offset(1, 0).EntireRow.Insert ' - insert new row
                            Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.Copy
                            k = k + 1
                            Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.PasteSpecial xlPasteFormats
                            Application.CutCopyMode = False
                            Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.Font.Bold = False
                            Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = Workbooks(Name).Worksheets(1).Cells(j, 4)
                            If Workbooks(Name).Worksheets(1).Cells(j, 6) = "" Then
                                Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks(Name).Worksheets(1).Cells(j, 8)
                            Else
                                Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks(Name).Worksheets(1).Cells(j, 6)
                            End If
                        End If
                        k = k + 1
                    Wend
                End If
                j = j + 1
            Wend
               
        ''Скандия
        ElseIf Workbooks(Name).Worksheets(1).Cells(j, 1) = "Раздел№1. Материалы и оборудование" Or Workbooks(Name).Worksheets(1).Cells(j, 1) = "Раздел №1. Материалы и оборудование" Then
            j = j + 1
            While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Итого" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> ""
                lastrow = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(Rows.Count, "A").End(xlUp).Row
                k = 3
                flag = False
                While flag = False ' - iteration between rows main pages
                    If Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = Workbooks(Name).Worksheets(1).Cells(j, 3) Then
                        flag = True
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2).Value + Workbooks(Name).Worksheets(1).Cells(j, 6)
        
                    ElseIf Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = "" Or k > lastrow Then
                        flag = True
                        k = lastrow
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).Offset(1, 0).EntireRow.Insert ' - insert new row
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.Copy
                        k = k + 1
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.PasteSpecial xlPasteFormats
                        Application.CutCopyMode = False
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1).EntireRow.Font.Bold = False
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 1) = Workbooks(Name).Worksheets(1).Cells(j, 3)
                        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(k, 2) = Workbooks(Name).Worksheets(1).Cells(j, 6)
                    End If
                    k = k + 1
                Wend
                j = j + 1
            Wend
        End If
       
        Workbooks("materials_loading.xlsm").Worksheets(1).Cells(1, 10) = Workbooks("materials_loading.xlsm").Worksheets(1).Cells(1, 10) + 1
        ''Закрытие файла
        Workbooks(Name).Close
    Next i
End If

End Sub



