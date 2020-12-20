Sub alphabetical()
    Dim ticker As String
    Dim year_change As Double
    Dim percent_change As Double
    Dim total_vol As LongLong
    Dim summary_table_row As Long
    Dim LastRow As Long
    Dim ws As Worksheet
   

      

    summary_table_row = 2
    start_row = 2
    
    For Each ws In Worksheets
    
    Range("j1").Value = "Ticker"
    Range("k1").Value = "Yearly Change"
    Range("l1").Value = "Percent Change"
    Range("m1").Value = "Total Volume"
   
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        For I = 2 To LastRow
            
            If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
                     start_value = ws.Cells(start_row, 3).Value
                     end_value = ws.Cells(I, 6).Value
                     ticker = ws.Cells(I, 1).Value
                     total_vol = total_vol + ws.Cells(I, 7).Value
                     year_change = end_value - start_value
                     percent_change = Round((year_change / start_value) * 100, 2)
                     ws.Range("J" & summary_table_row).Value = ticker
                     ws.Range("K" & summary_table_row).Value = year_change
                     ws.Range("L" & summary_table_row).Value = percent_change
                     ws.Range("M" & summary_table_row).Value = total_vol
                     summary_table_row = summary_table_row + 1
                     total_vol = 0
                
            Else:
            total_vol = total_vol + ws.Cells(I, 7).Value
            
            End If
            Next I
            
            
            
            For j = 2 To LastRow
            year_change = ws.Cells(j, 11).Value
            If year_change < 0 Then
            ws.Cells(j, 11).Interior.ColorIndex = 3
            Else
            ws.Cells(j, 11).Interior.ColorIndex = 4
            End If
            Next j
            
            Next ws
            
            
            
End Sub