Sub alphabetical_testing()
   
        Dim ticker_name As String
        Dim ticker As Double
        ticker = 0
        Dim ticker_row As Integer
        ticker_row = 2
        Dim open_price As Double
        open_price = Cells(2, 3).Value
        
        Dim close_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        last_row = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To last_row

         If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
      ticker_name = Cells(i, 1).Value
         ticker = ticker + Cells(i, 7).Value
        Range("I" & ticker_row).Value = ticker_name
    
        Range("L" & ticker_row).Value = ticker

       close_price = Cells(i, 6).Value
      yearly_change = (close_price - open_price)
        Range("J" & ticker_row).Value = yearly_change
      If (open_price = 0) Then
percent_change = 0
        Else
    percent_change = yearly_change / open_price
                
   End If


    Range("K" & ticker_row).Value = percent_change
    Range("K" & ticker_row).NumberFormat = "0.00%"
    ticker_row = ticker_row + 1
    ticker = 0
    open_price = Cells(i + 1, 3)
          
            Else
              ticker = ticker + Cells(i, 7).Value
            End If
    
    Next i
    last_row_summary_table = Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To last_row_summary_table
    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 10
   Else
    Cells(i, 10).Interior.ColorIndex = 3
    End If
  Next i
End Sub


