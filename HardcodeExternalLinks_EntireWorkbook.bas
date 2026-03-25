Attribute VB_Name = "HardcodeExternalLinks_EntireWorkbook"
Sub HardcodeExternalLinks_EntireWorkbook()

    Dim ws As Worksheet       ' Variable to hold each worksheet during the loop
    Dim rng As Range          ' Variable to hold the range of formula cells per sheet
    Dim cell As Range         ' Variable to iterate through each cell in the range
    Dim sheetCount As Long    ' Counter to track how many sheets were processed

    sheetCount = 0            ' Initialize the processed sheet counter

    Application.ScreenUpdating = False ' Disable screen refresh for better performance

    For Each ws In ThisWorkbook.Worksheets ' Loop through every worksheet in the workbook

        On Error Resume Next  ' Suppress errors in case no formula cells exist on this sheet
        Set rng = Nothing     ' Reset rng before each sheet to avoid carrying over previous reference
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas) ' Get only cells containing formulas
        On Error GoTo 0       ' Restore default error handling

        If Not rng Is Nothing Then ' Check if formula cells were found on this sheet

            For Each cell In rng ' Loop through each formula cell
                If InStr(1, cell.Formula, "[", vbTextCompare) > 0 Then ' Check if the formula contains "[" indicating an external link
                    cell.Value = cell.Value ' Replace the formula with its current calculated value
                End If
            Next cell         ' Move to the next cell

            sheetCount = sheetCount + 1 ' Increment the processed sheet counter

        End If

    Next ws                   ' Move to the next worksheet

    Application.ScreenUpdating = True ' Re-enable screen refresh

    MsgBox "External link formulas were converted to values across " & sheetCount & " sheet(s)." ' Notify the user with the number of sheets processed

End Sub
