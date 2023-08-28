Attribute VB_Name = "Module1"
'add four columns. Ticker, Yearly Change, Percent Change, Total Stock Volume
'Ticker column has ticker symbol
'Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year
'The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock. The result should match the following image:

Sub analysis():

'set up dimension
Dim total As Double
Dim i As Long
Dim change As Single
Dim j As Integer
Dim start As Long
Dim rowcount As Long
Dim percentChange As Single
Dim days As Integer
Dim dailyChange As Single
Dim averageChange As Single

'enter titles of new columns i-l
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Set initial values
j = 0
total = 0
change = 0
start = 2

'what is the row number of the last row with data
rowcount = Cells(Rows.Count, "A").End(xlUp).Row
For i = 2 To rowcount

'when ticker changes then give results
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Store results in variables
total = total + Cells(i, 7).Value

'if total volume is zero
If total = 0 Then
                
'print the results
Range("I" & 2 + j).Value = Cells(i, 1).Value
Range("J" & 2 + j).Value = 0
Range("K" & 2 + j).Value = "%" & 0
Range("L" & 2 + j).Value = 0

Else
 
'Find First non zero starting value
If Cells(start, 3) = 0 Then
For find_value = start To i
If Cells(find_value, 3).Value <> 0 Then
start = find_value

Exit For

End If

Next find_value


End If

'Calculate Change
change = (Cells(i, 6) - Cells(start, 3))

percentChange = change / Cells(start, 3)

'start of the next stock ticker
start = i + 1

'print the results
Range("I" & 2 + j).Value = Cells(i, 1).Value
Range("J" & 2 + j).Value = change
Range("J" & 2 + j).NumberFormat = "0.00"
Range("K" & 2 + j).Value = percentChange
Range("K" & 2 + j).NumberFormat = "0.00%"
Range("L" & 2 + j).Value = total

'shade positives green and negatives red
Select Case change
Case Is > 0
Range("J" & 2 + j).Interior.ColorIndex = 4
Case Is < 0
Range("J" & 2 + j).Interior.ColorIndex = 3
Case Else
Range("J" & 2 + j).Interior.ColorIndex = 0

End Select

End If

'reset variables for new stock ticker
 total = 0
 change = 0
 j = j + 1
 days = 0

'If ticker is still the same add results
Else: total = total + Cells(i, 7).Value

End If

Next i

End Sub

