VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Macros()

'variebles declaration
Dim numberSheets As Integer ' - actual amount of sheets
Dim lastrow As Long ' - actual amount of rows

'variebles initialization
numberSheets = ActiveWorkbook.Worksheets.Count ' - get actual amount of sheets


'Clearing old data
lastrow = Worksheets(1).Cells(rows.Count, "A").End(xlUp).Row
lastrow = lastrow - 10
For Row = 1 To lastrow
    If Worksheets(1).Cells(Row, 1) = "Всего за раздел" Then
        Worksheets(1).Cells(Row, 7) = ""
    ElseIf Worksheets(1).Cells(Row, 1) <> "Полы" And Worksheets(1).Cells(Row, 1) <> "Стены" And Worksheets(1).Cells(Row, 1) <> "Чистовая электрика" And Worksheets(1).Cells(Row, 1) <> "Двери" And Worksheets(1).Cells(Row, 1) <> "Оборудование" And Worksheets(1).Cells(Row, 1) <> "Мебель" And Worksheets(1).Cells(Row, 1) <> "Декор" And Worksheets(1).Cells(Row, 1) <> "Другое" And Worksheets(1).Cells(Row, 1) <> "Тип" Then
        For col = 2 To 7
            Worksheets(1).Cells(Row, col) = ""
        Next col
    End If
Next Row


'Uploading data from additional pages to main page
For i = 2 To numberSheets ' - iteration between worksheets
    lastrow = Worksheets(i).Cells(rows.Count, "A").End(xlUp).Row ' - get number of last row

    For j = 1 To lastrow ' - iteration between rows additional pages
        If Worksheets(i).Cells(j, 1) = "Полы" Or Worksheets(i).Cells(j, 1) = "Стены" Or Worksheets(i).Cells(j, 1) = "Чистовая электрика" Or Worksheets(i).Cells(j, 1) = "Двери" Or Worksheets(i).Cells(j, 1) = "Оборудование" Or Worksheets(i).Cells(j, 1) = "Мебель" Or Worksheets(i).Cells(j, 1) = "Декор" Or Worksheets(i).Cells(j, 1) = "Другое" Then
            topic = Worksheets(i).Cells(j, 1)
            
        ElseIf Worksheets(i).Cells(j, 1) = "Всего за раздел" Then
            k = 1
            While Worksheets(1).Cells(k, 1) <> topic
                k = k + 1
            Wend
            While Worksheets(1).Cells(k, 1) <> "Всего за раздел"
                k = k + 1
            Wend
            Worksheets(1).Cells(k, 7) = Worksheets(1).Cells(k, 7) + Worksheets(i).Cells(j, 7)
            
        ElseIf Worksheets(i).Cells(j, 1) <> "Тип" And Worksheets(i).Cells(j, 1) <> "" Then
            k = 1
            flag = False
            While flag = False ' - iteration between rows main pages
                If Worksheets(1).Cells(k, 1) = Worksheets(i).Cells(j, 1) Then
                    flag = True
                    For col = 2 To 7 ' - iteration between columns

                        If Worksheets(1).Cells(2, col) = "Количество" Or Worksheets(1).Cells(2, col) = "Стоимость" Then
                            Worksheets(1).Cells(k, col) = Worksheets(1).Cells(k, col) + Worksheets(i).Cells(j, col)

                        ElseIf Worksheets(1).Cells(2, col) = "Поставщик" Or Worksheets(1).Cells(2, col) = "Наименование" Or Worksheets(1).Cells(2, col) = "Артикул" Or Worksheets(1).Cells(2, col) = "Цена за ед." Then
                            If Worksheets(i).Cells(j, col) <> "" Then
                                Worksheets(1).Cells(k, col) = Worksheets(i).Cells(j, col)
                            End If
                        End If
                    Next col

                ElseIf Worksheets(1).Cells(k, 1) = "ИТОГО" Or k > 200 Then
                    flag = True
                    k = 1
                    While Worksheets(1).Cells(k, 1) <> topic
                        k = k + 1
                    Wend
                    While Worksheets(1).Cells(k, 1) <> "Всего за раздел"
                        k = k + 1
                    Wend
                    k = k - 1
                    Worksheets(1).Cells(k, 1).Offset(1, 0).EntireRow.Insert ' - insert new row
                    Worksheets(1).Cells(k, 1).EntireRow.Copy
                    k = k + 1
                    Worksheets(1).Cells(k, 1).EntireRow.PasteSpecial xlPasteFormats
                    Application.CutCopyMode = False
                    Worksheets(1).Cells(k, 1).EntireRow.Font.Bold = False
                    
                    For col = 1 To 7 ' - iteration between columns
                        If Worksheets(1).Cells(2, col) = "Количество" Or Worksheets(1).Cells(2, col) = "Стоимость" Then
                            Worksheets(1).Cells(k, col) = Worksheets(1).Cells(k, col) + Worksheets(i).Cells(j, col)

                        Else
                            If Worksheets(i).Cells(j, col) <> "" Then
                                Worksheets(1).Cells(k, col) = Worksheets(i).Cells(j, col)
                            End If
                        End If
                    Next col
                End If
                k = k + 1
            Wend
        End If
    Next j
Next i



End Sub

