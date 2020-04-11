Attribute VB_Name = "Module1"
Sub alphatest():

'dimension variable
Dim ticker As String
Dim yrchg As Double
Dim perchg As Double
Dim stockvol As Double
Dim initprice As Double
Dim ws As Worksheet

Dim summary_table_row As Integer

summary_table_row = 2

For Each ws In Worksheets
stockvol = 0
initprice = 0
finalprice = 0


Worksheets("A").Cells(1, 9).Value = "Ticker"
Worksheets("A").Cells(1, 10).Value = "Yearly Change"
Worksheets("A").Cells(1, 11).Value = "Percent Change"
Worksheets("A").Cells(1, 12).Value = "Total Stock Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'create array of ticker names and sum stock volume'

For i = 2 To LastRow
   
'start with initial value before loop'
  
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        finalprice = ws.Cells(i, 6).Value
       
        stockvol = stockvol + ws.Cells(i, 7).Value
        yrchg = initprice - finalprice
        If initprice = 0 Then
            prchg = "0.0"
            Else
            prchg = FormatPercent(yrchg / initprice)
        End If
        Sheets("A").Range("I" & summary_table_row).Value = ticker
        Sheets("A").Range("L" & summary_table_row).Value = stockvol
        Sheets("A").Range("J" & summary_table_row).Value = yrchg
        Sheets("A").Range("K" & summary_table_row).Value = prchg
        summary_table_row = summary_table_row + 1
        
        stockvol = 0
        initprice = 0
        yrchg = 0
        finalprice = 0
        
    
    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
        initprice = ws.Cells(i, 3).Value
        stockvol = stockvol + ws.Cells(i, 7).Value
        
    Else
        initprice = initprice
       
        
        End If
   
    Next i

For j = 2 To LastRow
    If Cells(j, 10) > 0 Then
        Cells(j, 10).Interior.ColorIndex = 4
    ElseIf Cells(j, 10) < 0 Then
        Cells(j, 10).Interior.ColorIndex = 3
    Else
        
    End If
Next j

Next ws

End Sub
