This is my code script 
Sub Report_Generator()
'This is to create the headers of the summary/report
'and we run the same code in each worksheet

For Each ws In Worksheets

ws.Range("L1").Value = "Ticker"
ws.Range("M1").Value = "Yearly Change"
ws.Range("N1").Value = "Percent Change"
ws.Range("O1").Value = "Total Stock Volume"
ws.Range("R2").Value = "Greatest % Increase"
ws.Range("R3").Value = "Greatest % Decrease"
ws.Range("R4").Value = "Greatest Total Volume"
ws.Range("S1").Value = "Ticker"
ws.Range("T1").Value = "Value"


'This is to fit the column width
ws.Columns("L:P").AutoFit
ws.Columns("R:T").AutoFit

'Define all the variables. For the yearly_change, we need to calculate the opening_price and closing_price
Dim ticker As String
Dim yearly_change As Double
Dim stock_volume As Double
Dim percent_change As Double
Dim opening_price As Double
Dim closing_price As Double
Dim row_number As Long
Dim last_row As Long
Dim last_row_2 As Long
Dim i As Long


'row_number is the ordinal number of row where we want to export the output, j starts as 2
row_number = 2
'stock_volume, opening_price equal zero before the calculation
stock_volume = 0
opening_price = 0
'last_row is the last row number that is non empty
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To last_row

'This is to stop the opening_price loop running
If opening_price = 0 Then
opening_price = ws.Cells(i, 3).Value
End If
 
  
'To compare if the next row has the same ticker on column A or cells(,1). If so, the total volume of that day is added to the stock_volume.
 If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
stock_volume = stock_volume + ws.Cells(i, 7).Value

 'to compare if the next row has the different ticker. If so, that would be the last addition to stock_volume, the closing_price would be the yearend closing price

    ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    stock_volume = stock_volume + ws.Cells(i, 7).Value
    closing_price = ws.Cells(i, 6).Value
    
'Export the sticker, the yearly change and total stock volume to column L,M and O respectively

    ws.Cells(row_number, 12).Value = ws.Cells(i, 1).Value
    ws.Cells(row_number, 13).Value = closing_price - opening_price
    
    'Format the color based on the yearly change value
    If ws.Cells(row_number, 13).Value < 0 Then
    ws.Cells(row_number, 13).Interior.Color = RGB(255, 0, 0)
    Else
    ws.Cells(row_number, 13).Interior.Color = RGB(0, 255, 0)
    End If
    
    ws.Cells(row_number, 14).Value = (closing_price - opening_price) / opening_price
    ws.Cells(row_number, 14).NumberFormat = "0.00%"
    
    ws.Cells(row_number, 15).Value = stock_volume
'Reset row_number to next ordinal number of row and stock_volume and opening_price back to zero for the new calculation
 row_number = row_number + 1
 stock_volume = 0
 opening_price = 0
 
End If

Next i
'To find the Greatest%Increase, %Decrease, and total volume
ws.Range("T2").Value = Application.WorksheetFunction.Max(ws.Range("N:N"))
ws.Range("T2").NumberFormat = "0.00%"
ws.Range("T3").Value = Application.WorksheetFunction.Min(ws.Range("N:N"))
ws.Range("T3").NumberFormat = "0.00%"
ws.Range("T4").Value = Application.WorksheetFunction.Max(ws.Range("O:O"))
ws.Range("T4").NumberFormat = "0.00E+00"

'last_row_2 is the last row on column N that is not empty. This is used for the if function
'to find the tickers that matche with values on column T
last_row_2 = ws.Cells(Rows.Count, 12).End(xlUp).Row

For i = 2 To last_row_2

If ws.Cells(i, 14).Value = ws.Range("T2").Value Then
ws.Range("S2").Value = ws.Cells(i, 12).Value

ElseIf ws.Cells(i, 14) = ws.Range("T3").Value Then
ws.Range("S3").Value = ws.Cells(i, 12).Value

ElseIf ws.Cells(i, 15) = ws.Range("T4").Value Then
ws.Range("S4").Value = ws.Cells(i, 12).Value

End If

Next i

Next ws

End Sub


