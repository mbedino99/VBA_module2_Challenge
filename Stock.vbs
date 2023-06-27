Sub moodule_2_challenge():

' create a script that loops through every row and outputs the following:
    ' ticker symbol
    ' yearly change from open to close
    ' percent change from open to close
    ' total stock volume of the stock
' for all worksheets
For Each ws In Worksheets


' define values
Dim ticker As String
Dim p_open As Double, p_close As Double
Dim column As Integer
Dim year_change As Double
Dim INC_max, DEC_max, VOL_max As Double
Dim INC_ticker, DEC_ticker, VOL_ticker As String


ws.Cells(1, "I").Value = "Ticker"
ws.Cells(1, "J").Value = "Yearly Change"
ws.Cells(1, "K").Value = "Percent Change"
ws.Cells(1, "L").Value = "Total Volume"
ws.Cells(2, "N").Value = "Greatest Percent Increase"
ws.Cells(3, "N").Value = "Greatest Percent Decrease"
ws.Cells(4, "N").Value = "Greatest Total Volume"
ws.Cells(1, "O").Value = "Ticker"
ws.Cells(1, "P").Value = "Value"


'find last row
lrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

' generate summary table for keeping track of results
Dim summary_table_row As Integer
summary_table_row = 2

'read in data:
vol_total = 0
column = 1
p_open = ws.Cells(2, 3).Value
INC_max = 0
DEC_max = 0
VOL_max = 0

For i = 2 To lrow
    ' search forchanges in ticker
    If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
    
    'set the ticker
    ticker = ws.Cells(i, 1).Value
    
    ' create yearly change
    p_close = ws.Cells(i, 6).Value
    ws.Range("j" & summary_table_row).Value = p_close - p_open
    If ws.Range("j" & summary_table_row).Value > 0 Then
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
    Else
    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
    End If
    
    'create percent change
    If p_open <> 0 Then
    ws.Range("K" & summary_table_row).Value = FormatPercent((p_close - p_open) / p_open, 2)
    Else
    ws.Range("K" & summary_table_row).Value = 0
    End If
    
    
    'add to the volume
    vol_total = vol_total + ws.Cells(i, 7).Value
    
    ' print ticker in summary table
    ws.Range("i" & summary_table_row).Value = ticker
    
    'print total volume on summary table
    ws.Range("l" & summary_table_row).Value = vol_total
    
    If ws.Range("K" & summary_table_row).Value > INC_max Then
    INC_max = ws.Range("K" & summary_table_row).Value
    INC_ticker = ws.Range("I" & summary_table_row).Value
    End If
    
    If ws.Range("K" & summary_table_row).Value < DEC_max Then
    DEC_max = ws.Range("K" & summary_table_row).Value
    DEC_ticker = ws.Range("I" & summary_table_row).Value
    End If
    
    If ws.Range("L" & summary_table_row).Value > VOL_max Then
    VOL_max = ws.Range("L" & summary_table_row).Value
    VOL_ticker = ws.Range("I" & summary_table_row).Value
    End If
    
    
    ' next line on summary table
    summary_table_row = summary_table_row + 1
    
    ' reser opening price
    p_open = ws.Cells(i + 1, "C").Value
    
    ' reset total volume
    vol_total = 0
    
    Else
    
    vol_total = vol_total + ws.Cells(i, 7).Value
    
    End If
    
Next i

ws.Cells(2, "P").Value = FormatPercent(INC_max, 2)
ws.Cells(2, "O").Value = INC_ticker

ws.Cells(3, "P").Value = FormatPercent(DEC_max, 2)
ws.Cells(3, "O").Value = DEC_ticker

ws.Cells(4, "P").Value = VOL_max
ws.Cells(4, "O").Value = VOL_ticker

' add fucntionality: "Greatest % increase", "greatest % decrease", and "greatest total volume"

' make it so it can run on every sheet in the workbook

' make it have conditional formatting to highlight + changein green and - change in red

Next ws

End Sub

