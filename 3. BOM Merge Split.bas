Option Explicit

' Main procedure to consolidate BOM by merging identical material numbers
Sub MergeBOM()
    Dim srcSheet As Worksheet, destSheet As Worksheet
    Dim lastRow As Long, i As Long, dict As Object
    Dim materialNum As String, newRow As Long
    
    ' Set source worksheet
    Set srcSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Create output sheet
    Set destSheet = CreateOutputSheet("Merged BOM")
    
    ' Copy header row
    srcSheet.Range("A1:G1").Copy destSheet.Range("A1")
    
    ' Find last data row
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Set dict = CreateObject("Scripting.Dictionary")
    newRow = 2
    
    For i = 2 To lastRow
        materialNum = Trim(srcSheet.Cells(i, 1).Value)
        
        If materialNum <> "" Then
            ' New material - add to dictionary
            If Not dict.Exists(materialNum) Then
                ' Store all column values as array
                ' A: Material Number
                ' B: Material Name
                ' C: Reference Designator(s)
                ' D: Package
                ' E: Mounting Type
                ' F: Quantity
                ' G: Unit
                dict.Add materialNum, Array( _
                    srcSheet.Cells(i, 1).Value, _
                    srcSheet.Cells(i, 2).Value, _
                    srcSheet.Cells(i, 3).Value, _
                    srcSheet.Cells(i, 4).Value, _
                    srcSheet.Cells(i, 5).Value, _
                    CDbl(srcSheet.Cells(i, 6).Value), _
                    srcSheet.Cells(i, 7).Value)
                
                ' Write to destination
                WriteRow destSheet, newRow, dict(materialNum)
                newRow = newRow + 1
            Else
                ' Existing material - merge data
                Dim existingData As Variant
                existingData = dict(materialNum)
                
                ' Append reference designator
                existingData(2) = existingData(2) & "," & srcSheet.Cells(i, 3).Value
                
                ' Sum quantities
                existingData(5) = existingData(5) + CDbl(srcSheet.Cells(i, 6).Value)
                
                ' Update dictionary
                dict(materialNum) = existingData
                
                ' Update destination row
                WriteRow destSheet, FindRowByMaterial(destSheet, materialNum), existingData
            End If
        End If
    Next i
    
    ' Apply final formatting
    ApplyFinalFormatting destSheet
    
    ' Completion message
    MsgBox "BOM consolidation complete!" & vbCrLf & _
           "Output: " & destSheet.Name, vbInformation
End Sub

' Main procedure to expand BOM by splitting combined reference designators
Sub SplitBOM()
    Dim srcSheet As Worksheet, destSheet As Worksheet
    Dim lastRow As Long, i As Long, j As Long
    Dim refDesignators() As String, refCount As Long, newRow As Long
    
    ' Set source worksheet
    Set srcSheet = ThisWorkbook.Sheets("Sheet1")
    
    ' Create output sheet
    Set destSheet = CreateOutputSheet("Split BOM")
    
    ' Copy header row
    srcSheet.Range("A1:G1").Copy destSheet.Range("A1")
    newRow = 2
    
    ' Find last data row
    lastRow = srcSheet.Cells(srcSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    For i = 2 To lastRow
        ' Split reference designators
        refDesignators = Split(Trim(srcSheet.Cells(i, 3).Value), ",")
        refCount = UBound(refDesignators) + 1
        
        ' Create new rows
        For j = 0 To UBound(refDesignators)
            ' Material Number
            destSheet.Cells(newRow, 1).Value = srcSheet.Cells(i, 1).Value
            ' Material Name
            destSheet.Cells(newRow, 2).Value = srcSheet.Cells(i, 2).Value
            ' Single Reference Designator
            destSheet.Cells(newRow, 3).Value = Trim(refDesignators(j))
            ' Package
            destSheet.Cells(newRow, 4).Value = srcSheet.Cells(i, 4).Value
            ' Mounting Type
            destSheet.Cells(newRow, 5).Value = srcSheet.Cells(i, 5).Value
            
            ' Quantity calculation
            If IsNumeric(srcSheet.Cells(i, 6).Value) Then
                destSheet.Cells(newRow, 6).Value = CDbl(srcSheet.Cells(i, 6).Value) / refCount
            Else
                destSheet.Cells(newRow, 6).Value = srcSheet.Cells(i, 6).Value
            End If
            
            ' Unit
            destSheet.Cells(newRow, 7).Value = srcSheet.Cells(i, 7).Value
            
            newRow = newRow + 1
        Next j
    Next i
    
    ' Apply final formatting
    ApplyFinalFormatting destSheet
    
    ' Completion message
    MsgBox "BOM expansion complete!" & vbCrLf & _
           "Output: " & destSheet.Name, vbInformation
End Sub

' Creates a new output worksheet
Function CreateOutputSheet(sheetName As String) As Worksheet
    ' Delete existing sheet if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create new sheet
    Set CreateOutputSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    CreateOutputSheet.Name = sheetName
End Function

' Writes data array to specified worksheet row
Sub WriteRow(ws As Worksheet, rowNum As Long, data As Variant)
    ' Validate row number
    If rowNum < 1 Then Exit Sub
    
    ' Write all columns
    ws.Cells(rowNum, 1).Value = data(0)  ' Material Number
    ws.Cells(rowNum, 2).Value = data(1)  ' Material Name
    ws.Cells(rowNum, 3).Value = data(2)  ' Reference Designator(s)
    ws.Cells(rowNum, 4).Value = data(3)  ' Package
    ws.Cells(rowNum, 5).Value = data(4)  ' Mounting Type
    ws.Cells(rowNum, 6).Value = data(5)  ' Quantity
    ws.Cells(rowNum, 7).Value = data(6)  ' Unit
End Sub

' Locates row number by material number
Function FindRowByMaterial(ws As Worksheet, material As String) As Long
    Dim lastRow As Long, i As Long
    
    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Search for material
    For i = 2 To lastRow
        If Trim(ws.Cells(i, 1).Value) = material Then
            FindRowByMaterial = i
            Exit Function
        End If
    Next i
    
    ' Not found
    FindRowByMaterial = 0
End Function

' Applies final formatting to worksheet
Sub ApplyFinalFormatting(ws As Worksheet)
    With ws
        ' Auto-fit columns
        .Columns.AutoFit
        
        ' Set column A to number format
        .Columns("A").NumberFormat = "0"
        
        ' Apply left alignment
        With .UsedRange
            .HorizontalAlignment = xlLeft
            ' Re-auto-fit after alignment
            .Columns.AutoFit
        End With
    End With
End Sub

