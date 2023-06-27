VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ticker():

Dim ws As Worksheet
For Each ws In Worksheets

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim sumrow As Integer
sumrow = 2

' Column creation
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Stock Total Volume"

' Output the ticker symbol

For i = 2 To lastrow
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
ws.Cells(sumrow, 9).Value = ws.Cells(i, 1).Value
sumrow = sumrow + 1
End If
Next i

Next ws
End Sub
 
' Calculate Yearly Change
Sub year_change():
For Each ws In Worksheets
Dim yearopen As Single
Dim yearclose As Single
Dim yearly_change As Single
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim sumrow As Integer
sumrow = 2

For i = 2 To lastrow
If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
yearopen = ws.Cells(i, 3).Value

ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
yearclose = ws.Cells(i, 6).Value

End If

' Finding yearly change and percent change output
If yearclose > 0 And yearopen > 0 Then
yearly_change = yearclose - yearopen
ws.Cells(sumrow, 10).Value = yearly_change

percent_change = (yearly_change / yearopen)
ws.Cells(sumrow, 11).Value = FormatPercent(percent_change)
sumrow = sumrow + 1
'reset
yearclose = 0
yearopen = 0
End If

' Conditional Formatting negative/positive change
If ws.Cells(i, 10).Value > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4
ElseIf ws.Cells(i, 10).Value < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3
End If
Next i

Next ws
End Sub
 
Sub stock_T():

For Each ws In Worksheets

Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim sumrow As Integer
sumrow = 2
Dim stocktotal As Double
stocktotal = 0

 ' code from class June 15
 For i = 2 To lastrow
    stocktotal = stocktotal + ws.Cells(i, 7).Value
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    ws.Cells(sumrow, 12).Value = stocktotal
    sumrow = sumrow + 1
    stocktotal = 0
     End If
 Next i
Next ws
End Sub

Sub Greatest():

For Each ws In Worksheets

' greatest decrease/increase/volume
'Column creation
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"
ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"

' assign variables

lastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

' Dim max As Single
Max = 0
maxticker = " "
' Dim min As Single
Min = 0
minticker = " "
' Greatest Total Stock Volume
maxtv = 0
maxtv_ticker = " "
' Loop through Yearly Change column to find min and max
For i = 2 To lastrow
  
    If ws.Cells(i, 11).Value > Max Then
    Max = ws.Cells(i, 11).Value
    maxticker = ws.Cells(i, 9).Value
   
    End If
   
    If ws.Cells(i, 11).Value < Min Then
    Min = ws.Cells(i, 11).Value
    minticker = ws.Cells(i, 9).Value
   
    End If
    
    ' find greatest total stock volume
    If ws.Cells(i, 12).Value > maxtv Then
    maxtv = ws.Cells(i, 12).Value
    maxtv_ticker = ws.Cells(i, 9).Value
    
    End If
    Next i

    ws.Cells(2, 15).Value = maxticker
    ws.Cells(2, 16).Value = FormatPercent(Max)
    ws.Cells(3, 15).Value = minticker
    ws.Cells(3, 16).Value = FormatPercent(Min)
    ws.Cells(4, 15).Value = maxtv_ticker
    ws.Cells(4, 16).Value = maxtv
    
Next ws

End Sub




