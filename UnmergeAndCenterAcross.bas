Attribute VB_Name = "UnmergeAndCenterAcross"
Sub UnmergeAndCenterAcross()
    Dim rng As Range
    Set rng = Selection

    With rng
        .UnMerge
        .HorizontalAlignment = xlCenterAcrossSelection
    End With
End Sub
