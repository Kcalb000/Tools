' Main procedure to compare data across three sheets (Sheet1, Sheet2, Sheet3)
' and export matching records to Sheet4 with source row highlighting
Sub CompareThreeSheets()
    ' Declare worksheet variables
    Dim ws1 As Worksheet, ws2 As Worksheet, ws3 As Worksheet, ws4 As Worksheet
    
    ' Declare row counters:
    Dim i As Long         ' Row index for Sheet1 (primary sheet)
    Dim j As Long         ' Row index for Sheet2 (first comparison sheet)
    Dim m As Long         ' Row index for Sheet3 (second comparison sheet)
    Dim k As Long         ' Output row counter for Sheet4 (results sheet)
    
    ' Declare match status flags:
    Dim matchFound1 As Boolean  ' Flag for E-value match in Sheet2
    Dim matchFound2 As Boolean  ' Flag for E-value match in Sheet3
    
    ' Set worksheet object references
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    Set ws3 = ThisWorkbook.Sheets("Sheet3")
    Set ws4 = ThisWorkbook.Sheets("Sheet4")
    
    ' Prepare Sheet4 for results output
    ws4.Cells.Clear  ' Clear all existing data
    ' Set comprehensive headers for output columns
    ws4.Range("A1:L1") = Array("Source", "Sheet1_C", "Sheet1_D", "Sheet1_E", _
                               "Sheet2_C", "Sheet2_D", "Sheet2_E", _
                               "Sheet3_C", "Sheet3_D", "Sheet3_E", _
                               "C_Equal", "Status")
    k = 1  ' Initialize output row counter (headers in row 1)
    
    ' Reset font color in all source sheets to black
    ' This clears any previous highlighting
    ws1.UsedRange.Font.Color = vbBlack
    ws2.UsedRange.Font.Color = vbBlack
    ws3.UsedRange.Font.Color = vbBlack
    
    ' Main processing loop: Rows 2-300 in Sheet1
    For i = 2 To 300
        ' Skip empty rows in Sheet1 column E
        If IsEmpty(ws1.Cells(i, "E")) Then GoTo NextRow
        
        ' Reset match flags for current row
        matchFound1 = False
        matchFound2 = False
        
        ' Search Sheet2 for matching E value
        For j = 2 To 300
            ' Skip empty rows in Sheet2
            If Not IsEmpty(ws2.Cells(j, "E")) Then
                ' Check for E column match between Sheet1 and Sheet2
                If ws1.Cells(i, "E") = ws2.Cells(j, "E") Then
                    matchFound1 = True  ' Set Sheet2 match flag
                    
                    ' Search Sheet3 for same E value
                    For m = 2 To 300
                        ' Skip empty rows in Sheet3
                        If Not IsEmpty(ws3.Cells(m, "E")) Then
                            ' Check for E column match between Sheet1 and Sheet3
                            If ws1.Cells(i, "E") = ws3.Cells(m, "E") Then
                                matchFound2 = True  ' Set Sheet3 match flag
                                
                                ' Compare C values across all three sheets
                                Dim cEqual As Boolean
                                cEqual = (ws1.Cells(i, "C") = ws2.Cells(j, "C") And _
                                         ws1.Cells(i, "C") = ws3.Cells(m, "C")
                                
                                ' Determine status message based on comparison
                                Dim status As String
                                If cEqual Then
                                    status = "All Equal"
                                Else
                                    status = "Not Equal"
                                End If
                                
                                ' Increment output counter
                                k = k + 1
                                
                                ' Write comprehensive results to Sheet4:
                                ' - Source information
                                ws4.Cells(k, "A") = "Sheet1 Row " & i
                                
                                ' - Sheet1 data (C, D, E columns)
                                ws4.Cells(k, "B") = ws1.Cells(i, "C")
                                ws4.Cells(k, "C") = ws1.Cells(i, "D")
                                ws4.Cells(k, "D") = ws1.Cells(i, "E")
                                
                                ' - Sheet2 data (C, D, E columns)
                                ws4.Cells(k, "E") = ws2.Cells(j, "C")
                                ws4.Cells(k, "F") = ws2.Cells(j, "D")
                                ws4.Cells(k, "G") = ws2.Cells(j, "E")
                                
                                ' - Sheet3 data (C, D, E columns)
                                ws4.Cells(k, "H") = ws3.Cells(m, "C")
                                ws4.Cells(k, "I") = ws3.Cells(m, "D")
                                ws4.Cells(k, "J") = ws3.Cells(m, "E")
                                
                                ' - Comparison results
                                ws4.Cells(k, "K") = cEqual  ' Boolean equality
                                ws4.Cells(k, "L") = status  ' Text status
                                
                                ' Highlight matching rows in red across all sheets
                                ws1.Rows(i).Font.Color = vbRed    ' Sheet1 row
                                ws2.Rows(j).Font.Color = vbRed    ' Sheet2 row
                                ws3.Rows(m).Font.Color = vbRed    ' Sheet3 row
                                
                                ' Exit Sheet3 loop after first match
                                Exit For
                            End If
                        End If
                    Next m  ' Continue to next row in Sheet3
                    
                    ' Exit Sheet2 loop after first match
                    Exit For
                End If
            End If
        Next j  ' Continue to next row in Sheet2
        
NextRow:
    Next i  ' Continue to next row in Sheet1
    
    ' Post-processing for Sheet4:
    ' Auto-resize columns for optimal readability
    ws4.Columns("A:L").AutoFit
    
    ' Format boolean column for clarity
    ws4.Range("K2:K" & k).NumberFormat = "TRUE;FALSE"
    
    ' Display completion message with statistics
    Dim recordCount As Long
    recordCount = k - 1  ' Calculate number of matching records
    MsgBox "Three-sheet comparison completed!" & vbCrLf & _
           "Sheet4 contains " & recordCount & " matching records" & vbCrLf & _
           "Matching rows highlighted in red in source sheets", _
           vbInformation, "Operation Complete"
End Sub
