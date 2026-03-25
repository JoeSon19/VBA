Attribute VB_Name = "HardcodeExternalLinks_ActiveSheet"
Sub HardcodeExternalLinks_ActiveSheet()

    Dim ws As Worksheet       ' Variable to hold the active worksheet reference
    Dim rng As Range          ' Variable to hold the range of formula cells
    Dim cell As Range         ' Variable to iterate through each cell in the range

    Set ws = ActiveSheet      ' Set ws to the currently active worksheet

    On Error Resume Next      ' Suppress errors in case no formula cells exist
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas) ' Get only cells containing formulas
    On Error GoTo 0           ' Restore default error handling

    If rng Is Nothing Then    ' Check if no formula cells were found
        MsgBox "No formula cells found on the active sheet."
        Exit Sub              ' Exit the macro early since there is nothing to process
    End If

    Application.ScreenUpdating = False ' Disable screen refresh for better performance

    For Each cell In rng      ' Loop through each formula cell
        If InStr(1, cell.Formula, "[", vbTextCompare) > 0 Then ' Check if the formula contains "[" indicating an external link
            cell.Value = cell.Value ' Replace the formula with its current calculated value
        End If
    Next cell                 ' Move to the next cell

    Application.ScreenUpdating = True ' Re-enable screen refresh

    MsgBox "External link formulas were converted to values on the active sheet." ' Notify the user that the process is complete

End Sub
