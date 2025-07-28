# Tools
For working application

# 1. BOM Compare.bas
  The macro processing be used for compare the standard 2 BOMs download for The TEAMCENTER.
  
  # how to use ?
    1. New create a Excel file, ensure the default sheet name is "Sheet1"
    2. Add 2 new sheets and the name shoud be "Sheet2" and "Sheet3"
    3. Copy all content of BOM 1 and BOM 2 with format into the new created Excel "Sheet1" and "Sheet2"
    4. Press the "Alt+F11" open the VBA edit page
    5. Select the VBAProject (Excal file name) then click the right key of mouse "import file" select the BOM Compare.bas then close the page
    6. Back to Excel file press "Alt+f8" run this "CompareSheets" macro processing
    
   # Key Components Explained:
Worksheet Setup

ws1: Primary data source (searched rows)

ws2: Comparison target (searched for matches)

ws3: Output destination for mismatched records

Processing Logic

Dual-loop structure: Outer loop (Sheet1 rows) vs Inner loop (Sheet2 rows)

E column matching: Primary comparison point

C column validation: Secondary comparison when E matches

Early exit: Stops searching Sheet2 after first match found

Output Handling

Sheet1 Status:

"Equal": E and C match between sheets

"Not Equal": E matches but C differs

"No Match": No corresponding E value found

Sheet3 Structure:

Columns A-C: Sheet1's C, D, E values

Columns D-F: Sheet2's C, D, E values

Row 1: Header labels

Dynamic row counter (k) tracks output position

Optimizations

Empty cell checks skip unnecessary comparisons

AutoFit ensures readable output in Sheet3

Clear separation of match handling logic

Workflow Summary:
Initialization

Clear Sheet3 and set headers

Initialize counters

Primary Processing

For each row in Sheet1 (2-300):

Search Sheet2 for matching E value

When found:

Compare C values

Mark status in Sheet1

Export mismatches to Sheet3

If no match found, mark "No Match"

Finalization

Format Sheet3 columns

Display completion message with record count

Important Notes:
Case Sensitivity

Comparisons use VBA's default case-insensitive comparison

For case-sensitive matching, use:

vba
If StrComp(ws1.Cells(i, "E"), ws2.Cells(j, "E"), vbBinaryCompare) = 0 Then
Performance Considerations

Processes 300x300 = 90,000 potential comparisons

Optimized with early exit and empty cell checks

For larger datasets, consider:

Using dictionary objects

Limiting search ranges dynamically

Error Handling

Add error checking at start:

vba
If ws1 Is Nothing Or ws2 Is Nothing Or ws3 Is Nothing Then
    MsgBox "Required sheets missing!", vbCritical
    Exit Sub
End If
This implementation provides comprehensive data comparison with clear result tracking and mismatch reporting while maintaining efficiency for the specified row range.
