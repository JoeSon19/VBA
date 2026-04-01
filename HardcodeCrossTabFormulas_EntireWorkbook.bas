Attribute VB_Name = "HardcodeCrossTabFormulas_EntireWorkbook"
Sub HardcodeCrossTabFormulas_EntireWorkbook()

    Dim ws As Worksheet       ' Variable to hold each worksheet during the loop
    Dim rng As Range          ' Variable to hold the range of formula cells per sheet
    Dim cell As Range         ' Variable to iterate through each cell in the range
    Dim convertedCount As Long ' Counter to track how many cells were converted

    convertedCount = 0        ' Initialize the converted cell counter

    Application.ScreenUpdating = False ' Disable screen refresh for better performance

    For Each ws In ThisWorkbook.Worksheets ' Loop through every worksheet in the workbook

        On Error Resume Next  ' Suppress errors in case no formula cells exist on this sheet
        Set rng = Nothing     ' Reset rng before each sheet to avoid carrying over previous reference
        Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas) ' Get only cells containing formulas
        On Error GoTo 0       ' Restore default error handling

        If Not rng Is Nothing Then ' Check if formula cells were found on this sheet

            For Each cell In rng ' Loop through each formula cell
                If InStr(1, cell.Formula, "!", vbTextCompare) > 0 Then ' Check if the formula contains "!" indicating a sheet reference
                    If InStr(1, cell.Formula, "[", vbTextCompare) = 0 Then ' Exclude external workbook links which contain "["
                        If Not IsError(cell.Value) Then ' Skip cells that evaluate to an error
                            If IsNumeric(cell.Value) Then ' Check if the result is a number
                                cell.Value = cell.Value ' Replace the formula with its current calculated value
                                convertedCount = convertedCount + 1 ' Increment the converted cell counter
                            End If
                        End If
                    End If
                End If
            Next cell         ' Move to the next cell

        End If

    Next ws                   ' Move to the next worksheet

    Application.ScreenUpdating = True ' Re-enable screen refresh

    MsgBox convertedCount & " cross-tab formula(s) converted to values across the workbook." ' Notify the user with the total converted count

End Sub
