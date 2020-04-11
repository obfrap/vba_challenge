Attribute VB_Name = "Module1"
Sub multiple_year_stock()

'Dimension variables

Dim ticker As String
Dim yrchng As Double
Dim perchng As Double
Dim stockvol As Double
Dim initprice As Double
Dim ws As Worksheet
Dim summary_table_row As Integer



'Create worksheet loop


For Each ws In Worksheets
summary_table_row = 2
stockvol = 0
initprice = 0
finalprice = 0

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To LastRow

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
        ws.Range("I" & summary_table_row).Value = ticker
        ws.Range("L" & summary_table_row).Value = stockvol
        ws.Range("J" & summary_table_row).Value = yrchg
        ws.Range("K" & summary_table_row).Value = prchg
        summary_table_row = summary_table_row + 1
        
        stockvol = 0
        initprice = 0
        yrchg = 0
        finalprice = 0
        
    
        ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value And ws.Cells(i - 1, 1) <> ws.Cells(i, 1) Then
            initprice = ws.Cells(i, 3).Value
            stockvol = stockvol + ws.Cells(i, 7).Value
        
    ElseIf ws.Cells(i + 1, 1) = "" Then

    Else
        initprice = initprice
               
        End If
   
    Next i

For j = 2 To LastRow
    If ws.Cells(j, 10) > 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 4
    ElseIf ws.Cells(j, 10) < 0 Then
        ws.Cells(j, 10).Interior.ColorIndex = 3
    Else
        
    End If
Next j

Next ws




End Sub
