' Main procedure to compare data between Sheet1 and Sheet2
' and export mismatched records to Sheet3
Sub CompareSheets()
    ' Declare worksheet variables for all sheets
    Dim ws1 As Worksheet  ' Primary comparison sheet
    Dim ws2 As Worksheet  ' Secondary comparison sheet
    Dim ws3 As Worksheet  ' Output sheet for mismatched records
    
    ' Declare loop counters and flags
    Dim i As Long         ' Row counter for Sheet1
    Dim j As Long         ' Row counter for Sheet2
    Dim k As Long         ' Output row counter for Sheet3
    Dim matchFound As Boolean  ' Flag indicating if match was found

    ' Set worksheet references by name
    Set ws1 = ThisWorkbook.Sheets("Sheet1")
    Set ws2 = ThisWorkbook.Sheets("Sheet2")
    Set ws3 = ThisWorkbook.Sheets("Sheet3")
    
    ' Prepare Sheet3 for output
    ws3.Cells.Clear      ' Clear all existing data
    ' Set column headers for output
    ws3.Range("A1:F1").Value = Array("Sheet1_C", "Sheet1_D", "Sheet1_E", "Sheet2_C", "Sheet2_D", "Sheet2_E")
    k = 1  ' Initialize output row counter (headers in row 1)
    
    ' Loop through Sheet1 rows 2-300 (row 1 typically contains headers)
    For i = 2 To 300
        matchFound = False  ' Reset match flag for each new row
        
        ' Process only if Sheet1 column E has data
        If Not IsEmpty(ws1.Cells(i, "E").Value) Then
            
            ' Loop through Sheet2 rows 2-300 to find matches
            For j = 2 To 300
                ' Process only if Sheet2 column E has data
                If Not IsEmpty(ws2.Cells(j, "E").Value) Then
                    
                    ' Check for value match in column E
                    If ws1.Cells(i, "E").Value = ws2.Cells(j, "E").Value Then
                        matchFound = True  ' Set match flag
                        
                        ' Compare values in column C
                        If ws1.Cells(i, "C").Value = ws2.Cells(j, "C").Value Then
                            ' Mark as equal in Sheet1 column F
                            ws1.Cells(i, "F").Value = "Equal"
                        Else
                            ' Mark as not equal in Sheet1 column F
                            ws1.Cells(i, "F").Value = "Not Equal"
                            
                            ' Increment output row counter
                            k = k + 1
                            
                            ' Export Sheet1 data to Sheet3
                            ws3.Cells(k, "A").Value = ws1.Cells(i, "C").Value  ' Sheet1 C
                            ws3.Cells(k, "B").Value = ws1.Cells(i, "D").Value  ' Sheet1 D
                            ws3.Cells(k, "C").Value = ws1.Cells(i, "E").Value  ' Sheet1 E
                            
                            ' Export Sheet2 data to Sheet3
                            ws3.Cells(k, "D").Value = ws2.Cells(j, "C").Value  ' Sheet2 C
                            ws3.Cells(k, "E").Value = ws2.Cells(j, "D").Value  ' Sheet2 D
                            ws3.Cells(k, "F").Value = ws2.Cells(j, "E").Value  ' Sheet2 E
                        End If
                        
                        ' Exit inner loop after first match found
                        Exit For
                    End If
                End If
            Next j  ' Continue to next row in Sheet2
        End If
        
        ' Handle case where no match was found in Sheet2
        If Not matchFound Then
            ws1.Cells(i, "F").Value = "No Match"
        End If
    Next i  ' Continue to next row in Sheet1
    
    ' Improve output formatting in Sheet3
    ws3.Columns("A:F").AutoFit  ' Resize columns to fit content
    
    ' Display completion message
    MsgBox "Data comparison completed!" & vbCrLf & _
           "Results: " & vbCrLf & _
           "• Sheet1 column F: Match status" & vbCrLf & _
           "• Sheet3: " & (k - 1) & " mismatched records", _
           vbInformation, "Operation Complete"
End Sub
