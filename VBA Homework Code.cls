VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockVBAHW_final()

For Each ws In Worksheets

    ws.Activate

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Columns("I:I").ColumnWidth = 19.17
    Columns("J:J").ColumnWidth = 19.67
    Columns("K:K").ColumnWidth = 20.83
    Columns("L:L").ColumnWidth = 22.67

Dim LastRow As Long
LastRow = Range("A" & Rows.Count).End(xlUp).Row
Dim i As Long
Dim Ticker_name As String
Dim Ticker_row As Long
Ticker_row = 2
Dim End_price As Double
End_price = 0
Dim Beg_price As Double
Beg_price = 0
Dim yearly_change As Double
yearly_change = 0
Dim percent_change As Double
percent_change = 0
Dim Total_volume As LongLong
Total_volume = 0

Beg_price = Cells(2, 3).Value

For i = 2 To LastRow
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Ticker_name = Cells(i, 1).Value
Range("I" & Ticker_row).Value = Ticker_name
End_price = Cells(i, 6).Value
yearly_change = End_price - Beg_price
Range("J" & Ticker_row).Value = yearly_change
        If yearly_change > 0 Then
                Cells(Ticker_row, 10).Interior.ColorIndex = 4
        ElseIf yearly_change < 0 Then
                Cells(Ticker_row, 10).Interior.ColorIndex = 3
        Else
                Cells(Ticker_row, 10).Interior.ColorIndex = 6
        End If
If Beg_price <> 0 Then
percent_change = yearly_change / Beg_price
Else
percent_change = 0
End If
Range("K" & Ticker_row).Value = percent_change
Cells(Ticker_row, 11).NumberFormat = "0.00%"
Total_volume = Total_volume + Cells(i, 7).Value
Range("L" & Ticker_row).Value = Total_volume
Ticker_row = Ticker_row + 1
Beg_price = Cells(i + 1, 3).Value

Else
Total_volume = Total_volume + Cells(i, 7).Value

End If

Next i

Next ws

End Sub
