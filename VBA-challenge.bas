Attribute VB_Name = "Module1"
Sub multipleyearstockdata()

    Dim ticker As String
    Dim summaryrow As Integer
    Dim i As Long
    Dim close_value As Double
    Dim open_value As Double
    Dim start As Long
    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim total_stock As Double
    Dim greatest_increase As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_total_volume As Double
    Dim greatest_total_volume_ticker As String

     
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        start = 2
        summaryrow = 2
        total_stock = 0
        greatest_increase = 0
        greatest_decrease = 0
        greatest_total_volume = 0
        
        ws.Cells(1, 8).Value = "Unique Ticker"
        ws.Cells(1, 9).Value = "Yearly change"
        ws.Cells(1, 10).Value = "Percentage change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
        
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            open_value = ws.Cells(start, 3).Value
            close_value = ws.Cells(i, 6).Value
            yearly_change = (close_value) - (open_value)
            If yearly_change < 0 Then
                ws.Cells(summaryrow, 9).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryrow, 9).Interior.ColorIndex = 4
            End If
                  
            percentage_change = (((close_value) - (open_value)) / (open_value))
            If percentage_change < 0 Then
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(summaryrow, 10).Interior.ColorIndex = 4
            End If
            
            If percentage_change > greatest_increase Then
                greatest_increase = percentage_change
                greatest_increase_ticker = ticker
            End If
            If percentage_change < greatest_decrease Then
                greatest_decrease = percentage_change
                greatest_decrease_ticker = ticker
            End If
           
     
            ws.Cells(summaryrow, 8).Value = ticker
            ws.Cells(summaryrow, 9).Value = yearly_change
            ws.Cells(summaryrow, 10).Value = percentage_change
            ws.Cells(summaryrow, 10).NumberFormat = "0.00%"
            ws.Cells(summaryrow, 11).Value = total_stock
            
           
           summaryrow = summaryrow + 1
            start = i + 1
            total_stock = 0
            Else:
                total_stock = total_stock + ws.Cells(i, 7).Value
                If total_stock > greatest_total_volume Then
                    greatest_total_volume = total_stock
                  greatest_total_volume_ticker = ticker
                End If
     
            End If
            
      
    
        Next i
        
            ws.Cells(2, 13).Value = "Greatest % increase"
            ws.Cells(3, 13).Value = "Greatest % decrease"
            ws.Cells(4, 13).Value = "Greatest total volume"
        
            ws.Cells(2, 14).Value = greatest_increase_ticker
            ws.Cells(2, 15).Value = greatest_increase
            ws.Cells(3, 14).Value = greatest_decrease_ticker
            ws.Cells(3, 15).Value = greatest_decrease
            ws.Cells(4, 14).Value = greatest_total_volume_ticker
            ws.Cells(4, 15).Value = greatest_total_volume
           
        
Next ws
        
End Sub
