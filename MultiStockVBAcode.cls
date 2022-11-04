VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub wth()

'identify the variables
'Dim ticker As String
'Dim vol As Long
'Dim i As Long
'Dim open_price As Double
'Dim close_price As Double
'Dim price_change As Double
'Dim price_change_percent As Double
'Dim stockvolume As Long
'Dim summary_table As Long
'Dim lr As Long

stockvolume = 0
summary_table = 2

'loop through the sheets
For Each ws In Worksheets

'label the headers
 
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

lr = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'let's get serious
        For i = 2 To lr
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                stockvolume = stockvolume + ws.Cells(i, 7).Value
                open_price = ws.Cells(i, 3).Value
                close_price = ws.Cells(i, 6).Value
            
                yearly_change = close_price - open_price
        
        'Put it in a chart
                ws.Cells(summary_table, 9).Value = ws.Cells(i, 1).Value
            
                ws.Cells(summary_table, 10).Value = yearly_change
            
                ws.Cells(summary_table, 11).Value = (yearly_change / open_price) * 100
            
                ws.Cells(summary_table, 12).Value = stockvolume
        
                If yearly_change > 0 Then
                    ws.Cells(summary_table, 10).Interior.ColorIndex = 4
                ElseIf yearly_change < 0 Then
                    ws.Cells(summary_table, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(summary_table, 10).Interior.ColorIndex = 2
                End If
                summary_table = summary_table + 1
                i = i + 1
            
        
            Else
                stockvolume = stockvolume + ws.Cells(i, 7).Value
        
            End If
        
        Next i

Next ws

End Sub




