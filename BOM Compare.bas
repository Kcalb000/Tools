' Main procedure to compare data between Sheet1 and Sheet2
' and export mismatched records to Sheet3
Sub CompareSheets()
    ' Declare worksheet variables
    Dim ws1 As Worksheet  ' Primary worksheet (source data)
    Dim ws2 As Worksheet  ' Comparison worksheet (target data)
    Dim ws3 As Worksheet  ' Output worksheet (mismatch results)
    
    ' Declare row counters
    Dim i As Long         ' Row index for Sheet1 (primary loop)
    Dim j As Long         ' Row index for Sheet2 (inner loop)
    Dim k As Long         ' Output row counter for Sheet3
    
    ' Declare match flag
    Dim matchFound As Boolean  ' Tracks if matching E value is found in Sheet2
    
    ' Set worksheet references
    Set ws1 = Sheets("Sheet1")
    Set ws2 = Sheets("Sheet2")
    Set ws3 = Sheets("Sheet3")
    
    ' Prepare Sheet3 for results
    ws3.Cells.Clear          ' Clear all existing data
    ' Set column headers
    ws3.Range("A1:F1") = Array("Sheet1_C", "Sheet1_D", "Sheet1_E", "Sheet2_C", "Sheet2_D", "Sheet2_E")
    k = 1  ' Initialize output row counter (headers in row 1)
    
    ' Main processing loop: Rows 2-300 in Sheet1
    For i = 2 To 300
        matchFound = False  ' Reset match flag for each new row
        
        ' Only process if column E has data
        If Not IsEmpty(ws1.Cells(i, "E")) Then
            
            ' Search Sheet2 for matching E value
            For j = 2 To 300
                ' Only process if Sheet2 column E has data
                If Not IsEmpty(ws2.Cells(j, "E")) Then
                    
                    ' Check for E column match
                    If ws1.Cells(i, "E") = ws2.Cells(j, "E") Then
                        matchFound = True  ' Set match flag
                        
                        ' Compare C column values
                        If ws1.Cells(i, "C") = ws2.Cells(j, "C") Then
                            ' Mark as equal in Sheet1 column F
                            ws1.Cells(i, "F") = "Equal"
                        Else
                            ' Mark as not equal in Sheet1 column F
                            ws1.Cells(i, "F") = "Not Equal"
                            
                            ' Increment output counter
                            k = k + 1
                            
                            ' Export Sheet1 data to Sheet3
                            ws3.Cells(k, "A") = ws1.Cells(i, "C")  ' C value
                            ws3.Cells(k, "B") = ws1.Cells(i, "D")  ' D value
                            ws3.Cells(k, "C") = ws1.Cells(i, "E")  ' E value
                            
                            ' Export Sheet2 data to Sheet3
                            ws3.Cells(k, "D") = ws2.Cells(j, "C")  ' C value
                            ws3.Cells(k, "E") = ws2.Cells(j, "D")  ' D value
                            ws3.Cells(k, "F") = ws2.Cells(j, "E")  ' E value
                        End If
                        
                        ' Exit inner loop after first match
                        Exit For
                    End If
                End If
            Next j  ' Continue to next row in Sheet2
        End If
        
        ' Handle no-match case
        If Not matchFound Then
            ' Mark as no match in Sheet1 column F
            ws1.Cells(i, "F") = "No Match"
        End If
    Next i  ' Continue to next row in Sheet1
    
    ' Format output in Sheet3
    ws3.Columns("A:F").AutoFit  ' Auto-resize columns for readability
    
    ' Completion message
    MsgBox "Task completed. Sheet3 has " & (k - 1) & " records."
End Sub
