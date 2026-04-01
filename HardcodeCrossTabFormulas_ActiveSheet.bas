Attribute VB_Name = "HardcodeCrossTabFormulas_ActiveSheet"
Sub HardcodeCrossTabFormulas_ActiveSheet()

    Dim ws As Worksheet       ' Variable to hold the active worksheet reference
    Dim rng As Range          ' Variable to hold the range of formula cells
    Dim cell As Range         ' Variable to iterate through each cell in the range
    Dim convertedCount As Long ' Counter to track how many cells were converted

    Set ws = ActiveSheet      ' Set ws to the currently active worksheet
    convertedCount = 0        ' Initialize the converted cell counter

    On Error Resume Next      ' Suppress errors in case no formula cells exist
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas) ' Get only cells containing formulas
    On Error GoTo 0           ' Restore default error handling

    If rng Is Nothing Then    ' Check if no formula cells were found
        MsgBox "No formula cells found on the active sheet."
        Exit Sub              ' Exit the macro early since there is nothing to process
    End If

    Application.ScreenUpdating = False ' Disable screen refresh for better performance

    For Each cell In rng      ' Loop through each formula cell
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
    Next cell                 ' Move to the next cell

    Application.ScreenUpdating = True ' Re-enable screen refresh

    MsgBox convertedCount & " cross-tab formula(s) converted to values on the active sheet." ' Notify the user with the total converted count

End Sub
