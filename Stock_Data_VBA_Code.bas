Attribute VB_Name = "Module1"
Sub stocks_final()
'Loop through worksheets
For Each ws In ActiveWorkbook.Worksheets

    'Define Variables
    Dim ticker_name As String
    
    Dim open_price As Double
    open_price = 0
    
    Dim close_price As Double
    close_price = 0
    
    Dim yearly_change As Double
    yearly_change = 0
    
    Dim stock_total As Double
    stock_total = 0
      
    Dim yearly_percent As Double
    yearly_percent = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Bonus Variables
    Dim max_ticker As String
    Dim min_ticker As String
    Dim max_value As Double
    max_value = 0
    Dim min_value As Double
    min_value = 0
    Dim greatest_total_vol As Double
    greatest_total_vol = 0
    Dim greatest_total_vol_name As String

    'Define first open price value
    open_price = ws.Cells(2, 3).Value
    
    'Headers for Bonus on each Worksheet
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
        
        'Loop through values for calculations/ find where ticker name changes
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker_name = ws.Cells(i, 1).Value
        
            stock_total = stock_total + ws.Cells(i, 7).Value
              
            close_price = ws.Cells(i, 6).Value
            yearly_change = close_price - open_price
            
            If open_price <> 0 Then
            yearly_percent = (yearly_change / open_price) * 100
            End If
            
            ws.Range("I" & Summary_Table_Row).Value = ticker_name
            ws.Range("J" & Summary_Table_Row).Value = yearly_change
            ws.Range("K" & Summary_Table_Row).Value = (CStr(yearly_percent) & "%")
            ws.Range("L" & Summary_Table_Row).Value = stock_total
            
            'conditional formatting
            If yearly_change > 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
            ElseIf yearly_change <= 0 Then
            ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If

            'bonus max/min calcs
            If yearly_percent < min_value Then
            min_value = yearly_percent
            min_ticker = ticker_name
            
            ElseIf yearly_percent > max_value Then
            max_value = yearly_percent
            max_ticker = ticker_name
            End If
            
            If stock_total > greatest_total_vol Then
            greatest_total_vol = stock_total
            greatest_total_vol_name = ticker_name
            End If
            
            'output bonus
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
            
            ws.Range("P2").Value = max_ticker
            ws.Range("P3").Value = min_ticker
            ws.Range("P4").Value = greatest_total_vol_name
            
            ws.Range("Q2").Value = max_value
            ws.Range("Q3").Value = min_value
            ws.Range("Q4").Value = greatest_total_vol
            
            
            'reset variables
            stock_total = 0
            Summary_Table_Row = Summary_Table_Row + 1
            open_price = ws.Cells(i + 1, 3).Value
            LastRow = 0

            Else
              stock_total = stock_total + ws.Cells(i, 7).Value
        End If
            
        Next i

Next ws

End Sub








