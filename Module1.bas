Attribute VB_Name = "Module1"
Sub CreateTextFile()
    
    ' Dependencies: Microsoft Word Object Library (Tools > References)
    Dim WordApp             As Object
    Dim fName               As String
    Dim LastRow             As Long
    Dim counter             As Integer
    Dim pathToDocx          As String
    Dim lanID               As String
    Dim brdName             As String
    Dim wbI                 As Workbook, wbO As Workbook
    Dim wsI                 As Worksheet
    Dim secondSheetSpot     As Integer
    Dim businessHeading     As String
    Dim extraSpot           As Integer
    Dim ws                  As Worksheet
    Dim FileExists          As Boolean
    Dim row                 As Integer
    Dim stringSpot          As Integer
    Dim cellString          As Integer
    
    With ThisWorkbook
    
        ' Checks for empties
        If Len(.Worksheets("Main").Range("D7").Value) = 0 Then
            .Worksheets("Main").Range("D7").Select
            MsgBox "Please enter BRD Filename!"
            Exit Sub
        End If
        If Len(.Worksheets("Main").Range("D8").Value) = 0 Then
            .Worksheets("Main").Range("D8").Select
            MsgBox "Please enter your LAN ID!"
            Exit Sub
        End If

        Application.StatusBar = "Starting process..."
        
        ' Creating new sheet
        Set ws = .Sheets.Add(After:= _
                 .Sheets(.Sheets.Count))
        ws.Name = "Temp"
        
        ' Setting filename and path for new text file
        lanID = LCase(.Worksheets("Main").Range("D8").Value)
        brdName = .Worksheets("Main").Range("D7").Value
        
    End With
    
    If Right(brdName, 5) <> ".docx" Then
        brdName = brdName & ".docx"
    End If
    
    ' Previously: ThisWorkbook.Path & "\" & "temp.txt"
    fName = "C:\Users\" & lanID & "\Desktop\" & "temp.txt "
    pathToDocx = "C:\Users\" & lanID & "\Desktop\" & brdName
       
    ' Checking if file exists
    If Len(Dir(pathToDocx)) = 0 Then
        With ThisWorkbook
            .Worksheets("Main").Activate
            Application.DisplayAlerts = False
            .Sheets("Temp").Delete
            Application.DisplayAlerts = True
            Application.StatusBar = ""
            MsgBox "Could not find BRD on Desktop. Please make sure filename is correct."
        End With
        Exit Sub
    End If
    
    ' Creating word object and opening BRD
    Set WordApp = CreateObject("Word.Application")
    WordApp.Documents.Open Filename:=pathToDocx, ReadOnly:=True ' before had only pathToDocx
    WordApp.Visible = True
    Application.StatusBar = "Progress: 25%, Opening BRD"
    Application.Wait (Now + TimeValue("0:00:05"))
    
    ' Saving BRD as text file
    WordApp.ActiveDocument.SaveAs2 Filename:=fName, _
        FileFormat:=2, _
        LockComments:=False, Password:="", AddToRecentFiles:=True, _
        WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts:=False, _
        SaveNativePictureFormat:=False, SaveFormsData:=False, SaveAsAOCELetter:= _
        False, Encoding:=1252, InsertLineBreaks:=False, AllowSubstitutions:=False, _
        LineEnding:=0, CompatibilityMode:=0

    Application.StatusBar = "Progress: 50%, Converting BRD to text file"
    Application.Wait (Now + TimeValue("0:00:05"))

    ' Closing documents and cleaning memory
    WordApp.Quit
    Set WordApp = Nothing
    
    ' Setting input and output worksheets and workbook
    Set wbI = ThisWorkbook
    Set wsI = wbI.Sheets("Temp")
    Set wbO = Workbooks.Open("C:\Users\" & lanID & "\Desktop\temp.txt")
    Application.StatusBar = "Progress: 60%, Opening temp text file"
    Application.Wait (Now + TimeValue("0:00:02"))
    
    ' Copying output text to input excel sheet
    wbO.Sheets(1).Cells.Copy wsI.Cells
    wbO.Close SaveChanges:=False

    ' Finding last used row
    With ThisWorkbook.Sheets("Temp")
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).row
    End With
    ThisWorkbook.Worksheets("Temp").Activate
    Application.StatusBar = "Progress: 65%, Copying contents over"
    Application.Wait (Now + TimeValue("0:00:01"))

    ' Instantiating variable for BR Heading and position in input sheet
    businessHeading = " "
    secondSheetSpot = 2
    
    Application.StatusBar = "Progress: 75%, Scraping contents for BR's and Headings"
    ' Looping thru Range of used cells and checking for requirements
    With ThisWorkbook
    
        For counter = 1 To LastRow
        
            ' Checking if first column has start of requirement
            If Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "BR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "BRL" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "NFR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "SR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "ACR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 4) = "DRBR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 4) = "PFCR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "HSR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "DMR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "IR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "CR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "TR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 3) = "DBR" Then
                
                ' Collecting data from first 3 cells if it fits criteria
                .Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value = Trim(CStr(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value) & _
                    " " & CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value) & _
                    " " & CStr(.Worksheets("Temp").Range("C" & CStr(counter)).Value)))
                
                ' Checking if there is extra information to add
                For extraSpot = 1 To LastRow - counter
                    
                    ' Checking if next row first column is a requirement or heading, if so then stop checking for extra info
                    If Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "Note" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 6) = "Source" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "BR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "BRL" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "NFR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "SR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "ACR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "DRBR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "PFCR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "HSR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "DMR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "IR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "CR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "TR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "5." Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "DBR" Then
                        
                        Exit For
                    
                    ' Checking if next row second column is a requirement or heading, if so then stop checking for extra info
                    ElseIf Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "Note" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 6) = "Source" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "BR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "BRL" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "NFR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "SR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "ACR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "DRBR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "PFCR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "HSR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "DMR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "IR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "CR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "TR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "5." Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "DBR" Then
                        
                        Exit For
                        
                    ' Checking for empty (roughly) row, if so then stop checking for extra info
                    ElseIf IsEmpty(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value))) = True And _
                        IsEmpty(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value))) = True And _
                        IsEmpty(Trim(CStr(.Worksheets("Temp").Range("C" & CStr(counter + extraSpot)).Value))) = True Then
                            
                        Exit For
                    
                    ' If none of exit scenarios were true, then adding extra info to BR cell
                    Else
                    
                        .Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value = Trim(CStr(CStr(.Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value) & Chr(10) & _
                            Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)) & " " & _
                            Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)) & " " & _
                            Trim(CStr(.Worksheets("Temp").Range("C" & CStr(counter + extraSpot)).Value))))
                    
                    End If
                Next
                
                secondSheetSpot = secondSheetSpot + 1
                
            ' Checking if second column has start of requirement
            ElseIf Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 2) = "BR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "BRL" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "NFR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 2) = "SR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "ACR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 4) = "DRBR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 4) = "PFCR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "HSR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "DMR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 2) = "IR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 2) = "CR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 2) = "TR" Or _
                Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value)), 3) = "DBR" Then
                
                ' Collecting data from next 3 cells if it fits criteria
                .Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value = Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value & _
                    " " & CStr(.Worksheets("Temp").Range("C" & CStr(counter)).Value) & _
                    " " & CStr(.Worksheets("Temp").Range("D" & CStr(counter)).Value)))
                
                ' Checking if there is extra information to add
                For extraSpot = 1 To LastRow - counter
                    
                    ' Checking if next row first column is a requirement or heading, if so then stop checking for extra info
                    If Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "Note" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 6) = "Source" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "BR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "BRL" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "NFR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "SR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "ACR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "DRBR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 4) = "PFCR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "HSR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "DMR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "IR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "CR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "TR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 2) = "5." Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)), 3) = "DBR" Then
                        
                        Exit For
                    
                    ' Checking if next row second column is a requirement or heading, if so then stop checking for extra info
                    ElseIf Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "Note" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 6) = "Source" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "BR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "BRL" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "NFR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "SR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "ACR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "DRBR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 4) = "PFCR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "HSR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "DMR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "IR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "CR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "TR" Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 2) = "5." Or _
                        Left(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)), 3) = "DBR" Then
                        
                        Exit For
                        
                    ' Checking for empty (roughly) row, if so then stop checking for extra info
                    ElseIf IsEmpty(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value))) = True And _
                        IsEmpty(Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value))) = True And _
                        IsEmpty(Trim(CStr(.Worksheets("Temp").Range("C" & CStr(counter + extraSpot)).Value))) = True And _
                        IsEmpty(Trim(CStr(.Worksheets("Temp").Range("D" & CStr(counter + extraSpot)).Value))) = True Then
                        
                        Exit For
    
                    ' If none of exit scenarios were true, then adding extra info to BR cell
                    Else
                    
                        .Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value = Trim(CStr(CStr(.Worksheets("Requirements").Range("C" & CStr(secondSheetSpot)).Value) & Chr(10) & _
                            Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter + extraSpot)).Value)) & " " & _
                            Trim(CStr(.Worksheets("Temp").Range("B" & CStr(counter + extraSpot)).Value)) & " " & _
                            Trim(CStr(.Worksheets("Temp").Range("C" & CStr(counter + extraSpot)).Value))))
                    
                    End If
                Next
                
                secondSheetSpot = secondSheetSpot + 1
                 
            ' Checking for BR Headings
            ElseIf Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "5." Or Left(Trim(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value)), 2) = "4." Then
                businessHeading = Trim(CStr(CStr(.Worksheets("Temp").Range("A" & CStr(counter)).Value) & _
                    " " & CStr(.Worksheets("Temp").Range("B" & CStr(counter)).Value) & _
                    " " & CStr(.Worksheets("Temp").Range("C" & CStr(counter)).Value)))
            End If
            
            ' Adding BR Heading
            .Worksheets("Requirements").Range("B" & CStr(secondSheetSpot)).Value = businessHeading
            .Worksheets("Requirements").Range("A" & CStr(secondSheetSpot)).Value = coverPageName
            
        Next counter
        
    
        Application.StatusBar = "Progress: 95%, Cleaning up"
        'Sheets(3).UsedRange.ClearContents
        ActiveWorkbook.Worksheets("Requirements").Activate
    
        ' Finding last row on newly updated sheet to format range
        With ThisWorkbook.Sheets("Requirements")
            LastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        End With
        
        .Worksheets("Requirements").Range("B2:C" & CStr(LastRow)).Borders.LineStyle = xlContinuous
        .Worksheets("Requirements").Range("B2:C" & CStr(LastRow)).WrapText = True
        
        ' DELETING TEMP FILE
        FileExists = (Dir(fName) <> "")
        If FileExists = True Then 'See above
          ' First remove readonly attribute, if set
          SetAttr fName, vbNormal
          ' Then delete the file
          Kill fName
        End If
    
        ' Trimming cells
        With ThisWorkbook.Worksheets("Requirements")
            For row = 2 To LastRow
                .Range("B" & row).Value = Trim(.Range("B" & row).Value)
                .Range("C" & row).Value = Trim(.Range("C" & row).Value)
                ' Trimming linebreaks
                stringSpot = Len(.Range("B" & row).Value)
                If Mid(.Range("B" & row).Value, stringSpot, 1) = Chr(10) _
                    Or Mid(.Range("B" & row).Value, stringSpot, 1) = Chr(13) _
                    Or Mid(.Range("B" & row).Value, stringSpot, 1) = vbCrLf Then

                    .Range("B" & row).Value = Mid(.Range("B" & row).Value, 1, stringSpot - 1)
                    Exit For

                End If
                tringSpot = Len(.Range("C" & row).Value)
                If Mid(.Range("C" & row).Value, stringSpot, 1) = Chr(10) _
                    Or Mid(.Range("C" & row).Value, stringSpot, 1) = Chr(13) _
                    Or Mid(.Range("C" & row).Value, stringSpot, 1) = vbCrLf Then

                    .Range("C" & row).Value = Mid(.Range("C" & row).Value, 1, stringSpot - 1)
                    Exit For

                End If
            Next
        End With
        
        Application.DisplayAlerts = False
        .Sheets("Temp").Delete
        Application.DisplayAlerts = True
        Application.StatusBar = ""
        MsgBox "Done!"
    
    End With

    
End Sub


Sub ClearRequirements()

    Dim LastRow As Integer

    With ThisWorkbook.Sheets("Requirements")
        LastRow = .Cells(.Rows.Count, "B").End(xlUp).row
        .Range("B2:C" & CStr(LastRow)).ClearContents
        .Range("B2:C" & CStr(LastRow)).Borders.LineStyle = xlLineStyleNone
        .Range("B1:C1").Borders.LineStyle = xlContinuous
    End With

End Sub



