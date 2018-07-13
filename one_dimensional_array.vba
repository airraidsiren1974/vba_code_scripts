'VBA Subroutine to automate the creation of a one-dimensional array - using a For next loop

Sub OneDimensionArray()

Dim curShippingCharges(5) As Currency
Dim iCounter As Integer

Worksheets("Sheet1").Activate

Range("B3").Activate

For iCounter = 0 To 5

    curShippingCharges(iCounter) = ActiveCell.Offset(iCounter, 1).Value
    
Next iCounter

'For iCounter = 0 To 5 Step 2
For iCounter = 5 To 0 Step -2

    MsgBox (curShippingCharges(iCounter))

Next iCounter

End Sub
