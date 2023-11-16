Sub analysis()

' Define variables
Dim rowcount As Long
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim yearchange As Double
Dim perchange As Double
Dim totvol As Double
Dim sumrow As Integer
Dim ws As Worksheet
Dim ginc As Double
Dim gdec As Double
Dim gvol As Double
Dim ginct As String
Dim gdect As String
Dim gvolt As String

' Iterate through each worksheet
For Each ws In Worksheets

' Reset variables for worksheet
sumrow = 2
ginc = 0
gdec = 0
gvol = 0
rowcount = ws.UsedRange.Rows.count

' Print headers for summary columns
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly Change"
ws.Cells(1, 11) = "Percent Change"
ws.Cells(1, 12) = "Total Volume"

' Iterate through all the rows in the sheet
For r = 2 To rowcount
ticker = ws.Cells(r, 1) ' Get ticker value for current row
totvol = totvol + ws.Cells(r, 7) ' Add current row to the total volume
If ws.Cells(r, 1) <> ws.Cells(r - 1, 1) Then ' If row is a new ticker then set the open price
openprice = ws.Cells(r, 3)
End If
If ws.Cells(r, 1) <> ws.Cells(r + 1, 1) Then ' Decide if the next row is a new ticker
closeprice = ws.Cells(r, 6)
yearchange = closeprice - openprice
perchange = yearchange / openprice

' Print summary values for the ticker
ws.Cells(sumrow, 9) = ticker
ws.Cells(sumrow, 10) = yearchange
ws.Cells(sumrow, 11) = perchange
ws.Cells(sumrow, 12) = totvol

ws.Cells(sumrow, 10).NumberFormat = "0.00"
ws.Cells(sumrow, 11).NumberFormat = "0.00%"

' Conditional formatting on yearchange cell
If yearchange < 0 Then
    ws.Cells(sumrow, 10).Interior.ColorIndex = 3
ElseIf yearchange > 0 Then
    ws.Cells(sumrow, 10).Interior.ColorIndex = 4
End If
sumrow = sumrow + 1
totvol = 0
End If

Next r

' Find greatest increase, decrease, and volume
For s = 2 To sumrow
If ws.Cells(s, 11) > ginc Then
    ginc = ws.Cells(s, 11)
    ginct = ws.Cells(s, 9)
ElseIf ws.Cells(s, 11) < gdec Then
    gdec = ws.Cells(s, 11)
    gdect = ws.Cells(s, 9)
End If
If ws.Cells(s, 12) > gvol Then
    gvol = ws.Cells(s, 12)
    gvolt = ws.Cells(s, 9)
End If
Next s

' Print summary of greatest increase, decrease, and volume
ws.Cells(1, 16) = "Ticker"
ws.Cells(1, 17) = "Value"
ws.Cells(2, 15) = "Greatest % Increase"
ws.Cells(2, 16) = ginct
ws.Cells(2, 17) = ginc
ws.Cells(3, 15) = "Greatest % Decrease"
ws.Cells(3, 16) = gdect
ws.Cells(3, 17) = gdec
ws.Cells(4, 15) = "Greatest Volume"
ws.Cells(4, 16) = gvolt
ws.Cells(4, 17) = gvol

ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 17).NumberFormat = "0"
' Autofit worksheet summary columns
ws.Columns("I:L").AutoFit
ws.Columns("O:Q").AutoFit
Next ws
End Sub