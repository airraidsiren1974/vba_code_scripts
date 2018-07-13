'VBA course on Lynda/Linkedin Learning
'For Each Loop - Stepping through all items in a collection

Sub CitiesArray()

Dim strCities(3) As String
Dim var As Variant

strCities(0) = "Los Angeles"
strCities(1) = "Portland"
strCities(2) = "Seattle"
strCities(3) = "DC"

For Each var In strCities

    MsgBox (var)

Next var

End Sub

Sub WorksheetNames()

Dim wbk As Workbook
Dim wks As Worksheet

Set wbk = ThisWorkbook

For Each wks In wbk.Worksheets

    wks.Name = wks.Name & "Test"

Next wks

End Sub
