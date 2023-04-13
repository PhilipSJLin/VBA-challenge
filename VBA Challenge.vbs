Sub VBAchallenge()
    Dim ws As Worksheet
    
   ' Loop through all sheets
    For Each ws In Worksheets
    
    ' Set an intial variable for holding the ticker name
    Dim Ticker As String
    Dim New_Ticker As String
    
    ' Set an initial variable for holding the yearly change
    Dim Yearly_Change As Double
    Dim Last_Close As Double
    Dim First_Open As Double
    Dim Number_of_Rows As LongLong
    
    ' Set an initial variable for holding the percent change
    Dim Percent_Change As Double
    
    ' Set an initial variable for holding the total stock volume
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
    
    ' Determine the Last Row
    ' LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
'---------------------------------------------------------------------------
    
        ' Loop through all the tickers
        Dim i As LongLong
        For i = 2 To LastRow
        
            ' Check if we are still within the same ticker.
            Ticker = ws.Cells(i, 1).Value
            New_Ticker = ws.Cells(i + 1, 1).Value
        
            If New_Ticker <> Ticker Then
        
                ' Print the ticker in the Summary Table
                ws.Cells(Summary_Table_Row, 9).Value = Ticker
            
                Number_of_Rows = Application.CountIf(Range("A:A"), Ticker)
            
                ' Calculate Yearly Change
                Last_Close = ws.Cells(i, 6).Value
                First_Open = ws.Cells(i - Number_of_Rows + 1, 3).Value
                Yearly_Change = Last_Close - First_Open
            
                ' Print the Yearly Change to the Summary Table
                ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
                
                ' Calculate Percent Change
                Percent_Change = (Last_Close - First_Open) / First_Open
                
                ' Print the Percent Change to the Summary Table
                ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
                ws.Range("K:K").NumberFormat = "0.00%"
            
            
                ' Calculate Total Stock Volume
                 Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                 
                ' Print the Total Stock Volume to the Summary Table
                ws.Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            ' Reset the Values
            Yearly_Change = 0
            Percent_Change = 0
            Total_Stock_Volume = 0

            Else

            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
             
            End If
        
        Next i
                
'---------------------------------------------------------------------------

    Dim j As LongLong
    
   'Determine the Last Row in Yearly Change column
    LastRow_YearlyChange = Cells(Rows.Count, 10).End(xlUp).Row
    
        'Loop through all the Yearly Change
        For j = 2 To LastRow_YearlyChange
        
            'Compare Yearly Change values
            If Cells(j, 10).Value >= 0 Then
            
            ws.Cells(j, 10).Interior.ColorIndex = 4

            Else
            
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
        
        Next j

'---------------------------------------------------------------------------

'Set an initial variable for holding the total stock volume
    Dim Max As Double
    Dim Min As Double
    Dim MaxVolume As LongLong
    Dim p As Integer
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    
    
    ' Determine the Last Rows in each columns in summary table
    LastRow_TickerInSummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    LastRow_PercentChange = ws.Cells(Rows.Count, 11).End(xlUp).Row
    LastRow_TotalStockVolume = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    'Find the Greatest percentage Increase, Greatest percentage Decrease, Total Stock Volume values
    Max = WorksheetFunction.Max(ws.Range("K2:K" & LastRow_PercentChange))
    Min = WorksheetFunction.Min(ws.Range("K2:K" & LastRow_PercentChange))
    MaxVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow_TotalStockVolume))
    
    'Print values
    ws.Cells(2, 17).Value = Max
    ws.Cells(3, 17).Value = Min
    ws.Cells(4, 17).Value = MaxVolume
    
    'Loop through each of the ticker in Summary table
    For p = 2 To LastRow_TickerInSummary
    
    If ws.Cells(p, 11).Value = Max Then
    ws.Cells(2, 16).Value = ws.Cells(p, 9)
    
    ElseIf ws.Cells(p, 11).Value = Min Then
    ws.Cells(3, 16).Value = ws.Cells(p, 9)
    
    ElseIf ws.Cells(p, 12).Value = MaxVolume Then
    ws.Cells(4, 16).Value = ws.Cells(p, 9)
    
    End If
    
    Next p

'Autofit Columns
ws.Columns("A:Q").AutoFit

Next ws

End Sub
