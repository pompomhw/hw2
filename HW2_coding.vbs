Attribute VB_Name = "Module1"
Sub Stock_Market_Analysis()

Dim ws As Worksheet
Dim yearly_change As Single
Dim t_lastrow, e_lastrow, ticker_rank, i, m As Integer
't_lastrow   : lastrow overall
'e_lastrow   : lastrow for each indiviual ticker
'ticker_rank : order of each ticker among the tickers in the sheet

 For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    t_lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ticker_rank = 2
    t_vol = 0
    e_firstrow = 2
    
    For i = 2 To t_lastrow
     If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
       t_vol = t_vol + ws.Cells(i, 7).Value
       
       first_open_price = ws.Cells(e_firstrow, 3).Value
       last_close_price = ws.Cells(i, 6).Value
       yearly_change = last_close_price - first_open_price
       
       '''' defining percent_change in case of the zero-denominator''''''''
       If first_open_price <> 0 Then
        percent_change = yearly_change / first_open_price
       Else
        percent_change = 0
       End If
           
       ws.Range("I" & ticker_rank).Value = ws.Cells(i, 1).Value
       ws.Range("j" & ticker_rank).Value = yearly_change
       ws.Range("K" & ticker_rank).Value = percent_change
       ws.Range("K" & ticker_rank).Style = "Percent"
       ws.Range("K" & ticker_rank).NumberFormat = "0.00%"
       ws.Range("L" & ticker_rank).Value = t_vol
       
       ''''' coloring the cells ''''''''''''''''''''''''''''''''''''''''''''
       If ws.Range("j" & ticker_rank).Value > 0 Then
         ws.Range("j" & ticker_rank).Interior.ColorIndex = 43
       ElseIf ws.Range("j" & ticker_rank).Value < 0 Then
         ws.Range("j" & ticker_rank).Interior.ColorIndex = 46
       End If
         
       
       e_firstrow = i + 1
       ticker_rank = ticker_rank + 1
       t_vol = 0
      
     Else
       t_vol = t_vol + ws.Cells(i, 7).Value
     End If
       
    Next i
    
    
    '''''''''''''hard addition'''''''''''''''''''''''''''''''''
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("p1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    lastrow = ws.Cells(Rows.Count, 10).End(xlUp).Row
    max_change = WorksheetFunction.Max(ws.Range("k2:" & "k" & lastrow))
    min_change = WorksheetFunction.Min(ws.Range("k2:" & "k" & lastrow))
    max_vol = WorksheetFunction.Max(ws.Range("L2:" & "L" & lastrow))
    
    Match_max_change = WorksheetFunction.Match(max_change, ws.Range("k2:" & "k" & lastrow), 0)
    match_min_change = WorksheetFunction.Match(min_change, ws.Range("k2:" & "k" & lastrow), 0)
    match_max_vol = WorksheetFunction.Match(max_vol, ws.Range("L2:" & "L" & lastrow), 0)
    
    ws.Range("p2").Value = ws.Cells(Match_max_change + 1, 9).Value
    ws.Range("p3").Value = ws.Cells(match_min_change + 1, 9).Value
    ws.Range("p4").Value = ws.Cells(match_max_vol + 1, 9).Value
    
    ws.Range("q2").Value = max_change
    ws.Range("q2").Style = "percent"
    ws.Range("q2").NumberFormat = "0.00%"
    ws.Range("q3").Value = min_change
    ws.Range("q3").Style = "percent"
    ws.Range("q3").NumberFormat = "0.00%"
    ws.Range("q4").Value = max_vol

 Next ws
 
End Sub




