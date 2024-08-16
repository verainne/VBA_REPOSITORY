# VBA_REPOSITORY

Sub tickerloop()
'Loop through worksheets
    For Each ws In Worksheets

        'Ticker name variable
        Dim ticker As String
    
        'Volume Variable
        Dim volume As Double
        volume = 0

        'Location Tracker
        Dim summary_row As Long
        summary_row = 2
        
        'Quarterly change opening price starting thingie
        Dim opening_price As Double
        'Set opening price
        opening_price = Cells(2, 3).Value
        
        'setting some other important variables
        Dim closing_price As Double
        Dim quarterly_changes As Double
        Dim percent_change As Double

        'Summary Table Names
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Quarterly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

        'Last Row code
    Dim LastRow As Long
  LastRow = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row

        'Loop through rows by using tickers
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
              ticker = Cells(i, 1).Value
              volume = volume + Cells(i, 7).Value
              Range("I" & summary_row).Value = ticker
              Range("L" & summary_row).Value = volume
        'the codes needed for quarterly difference
              closing_price = Cells(i, 6).Value
              quarterly_changes = (closing_price - opening_price)
              Range("J" & summary_row).Value = quarterly_changes
                If (opening_price = 0) Then
                    percentage = 0

                Else
                    
                    percentage = quarterly_changes / opening_price
                
                End If
              Range("K" & summary_row).Value = percentage
              Range("K" & summary_row).NumberFormat = "0.00%"
              summary_row = summary_row + 1
              volume = 0
              opening_price = Cells(i + 1, 3)
            
            Else
              volume = volume + Cells(i, 7).Value
            
            End If
        
        Next i
    
    'Color coding
    
    For i = 2 To LastRow
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
    Next i
    
    Next ws

End Sub
