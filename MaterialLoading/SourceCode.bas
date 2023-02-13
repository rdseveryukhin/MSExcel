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

'' Path to directory
myPath = Range("B1")

Set fso = CreateObject("Scripting.FileSystemObject")
Set myFolder = fso.GetFolder(myPath)

'' If directory is empty
If myFolder.Files.Count = 0 Then
    MsgBox "No files"
    Exit Sub
Else
    Dim count_files As Integer
    count_files = myFolder.Files.Count
    
    '' Array with file names
    ReDim myFiles(1 To count_files)

        '' Filling array with paths to files
    For Each myFile In myFolder.Files
        i = i + 1
        myFiles(i) = myFile.Path
    Next
    
    '' Processing
    For i = 1 To count_files
        Workbooks.Open (myFiles(i))
        Name = Right(myFiles(i), Len(myFiles(i)) - InStrRev(myFiles(i), "\", , 1))
        
    ''Searching title
        j = 1
        While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Ðàçäåë 2. Ìàòåðèàëû è îáîðóäîâàíèå â òåêóùèõ öåíàõ" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Ðàçäåë¹1. Ìàòåðèàëû è îáîðóäîâàíèå" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Ðàçäåë ¹1. Ìàòåðèàëû è îáîðóäîâàíèå"
            j = j + 1
        Wend
        
        ''Andersen
        If Workbooks(Name).Worksheets(1).Cells(j, 1) = "Ðàçäåë 2. Ìàòåðèàëû è îáîðóäîâàíèå â òåêóùèõ öåíàõ" Then
            j = j + 1
            While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Èòîãè ïî àêòó:" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Èòîãî ïî ðàçäåëó 2 Ìàòåðèàëû è îáîðóäîâàíèå â òåêóùèõ öåíàõ" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "ÈÒÎÃÈ ÏÎ ÀÊÒÓ:"
                If Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Íèæíèé ÿðóñ" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Ñåêöèÿ 5.6" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Ñåêöèÿ 5.5" Then
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
               
            ''Scandy
        ElseIf Workbooks(Name).Worksheets(1).Cells(j, 1) = "Ðàçäåë¹1. Ìàòåðèàëû è îáîðóäîâàíèå" Or Workbooks(Name).Worksheets(1).Cells(j, 1) = "Ðàçäåë ¹1. Ìàòåðèàëû è îáîðóäîâàíèå" Then
            j = j + 1
            While Workbooks(Name).Worksheets(1).Cells(j, 1) <> "Èòîãî" And Workbooks(Name).Worksheets(1).Cells(j, 1) <> ""
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
        ''Closing file
        Workbooks(Name).Close
    Next i
End If

End Sub



