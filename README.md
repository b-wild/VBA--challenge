# VBA--challenge
week 2 challenge

created code to list securities and total stock volumn in the test ABC file

For i = 2 To LastRow
    If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        ticker = Cells(i, 1).Value
        volumn = volumn + Cells(i, 7).Value
    
        Range("I" & Summary_Table_Row).Value = ticker
        Range("L" & Summary_Table_Row).Value = volumn
        'adding row in table
        Summary_Table_Row = Summary_Table_Row + 1
        'volumn reset
        volumn = 0
    Else
        volumn = volumn + Cells(i, 7).Value
End If
