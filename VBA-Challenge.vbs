Sub VBA_challenge()

'screen for worksheet

  Dim WS_Count As Integer
  Dim k As Integer


  ' Set WS_Count equal to the number of worksheets in the active
  ' workbook.
  WS_Count = ActiveWorkbook.Worksheets.Count

  ' Begin the loop.
  For k = 1 To WS_Count

        Worksheets(k).Select

        'set variable
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_volume As LongLong
        Dim Greatest_increase As Double
        Dim Greatest_decrease As Double
        Dim Greatest_volume As LongLong
        Dim increase_ticker As String
        Dim decrease_ticker As String
        Dim volume_ticker As String
        
        
        
        yearly_change = 0
        total_volume = 0
        input_row = 2
        Greatest_increase = 0
        Greatest_decrease = 0
        Greatest_volume = 0
        
        
        'add header
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        
        
        'find out last row
         LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        

'screen for all row
    For I = 2 To LastRow
    
        If Cells(I, 1).Value = Cells(I + 1, 1).Value Then
        total_volume = total_volume + Cells(I, 7).Value
        yearly_change = yearly_change + Cells(I + 1, 3).Value - Cells(I, 3).Value
        
    
        ElseIf Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        total_volume = total_volume + Cells(I, 7).Value
        percent_change = (yearly_change - Cells(I, 3).Value + Cells(I, 6).Value) / -(yearly_change - Cells(I, 3).Value)
        yearly_change = yearly_change - Cells(I, 3).Value + Cells(I, 6).Value
        
        
        Cells(input_row, 10).Value = yearly_change
        Cells(input_row, 12).Value = total_volume
        Cells(input_row, 9).Value = Cells(I, 1)
        Cells(input_row, 11).Value = percent_change
        Cells(input_row, 11).NumberFormat = "0.00%"
        
            If Cells(input_row, 10).Value > 0 Then
                Cells(input_row, 10).Interior.ColorIndex = 4
            ElseIf Cells(input_row, 10).Value < 0 Then
                Cells(input_row, 10).Interior.ColorIndex = 3
                
            End If
            

             
            
            
            
        
        total_volume = 0
        yearly_change = 0
        input_row = input_row + 1
        
        
       
        
        End If
        
    Next I
    
    'Bonus------------------------------------------------------------------
    
    'add Header
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    
    
        For a = 2 To LastRow
        
        
            'Bonus: find out greatest increase
            If Cells(a, 11).Value >= Greatest_increase Then
                Greatest_increase = Cells(a, 11).Value
                increase_ticker = Cells(a, 9).Value
                
            ElseIf Cells(a, 11).Value < Greatest_increase Then
                Greatest_increase = Greatest_increase
                increase_ticker = increase_ticker
            
            End If
       
             
        Next a
        
             Cells(2, 17).Value = Greatest_increase
             Cells(2, 17).NumberFormat = "0.00%"
             Range("P2").Value = increase_ticker
             
        
            
             
             
             
             
        For b = 2 To LastRow
        
            'Bonus: Find out greatest decrease
            If Cells(b, 11).Value <= Greatest_decrease Then
                Greatest_decrease = Cells(b, 11).Value
                decrease_ticker = Cells(b, 9).Value
            
                
            ElseIf Cells(b, 11).Value > Greatest_decrease Then
                Greatest_decrease = Greatest_decrease
                decrease_ticker = decrease_ticker
                
            End If

    Next b
            Cells(3, 17).Value = Greatest_decrease
            Cells(3, 17).NumberFormat = "0.00%"
            Cells(3, 16).Value = decrease_ticker
            
        
        For c = 2 To LastRow
        
            If Cells(c, 12).Value >= Greatest_volume Then
                Greatest_volume = Cells(c, 12).Value
                volume_ticker = Cells(c, 9).Value
                
            ElseIf Cells(c, 12).Value < Greatest_volume Then
                Greatest_volume = Greatest_volume
                volume_ticker = volume_ticker
                
            End If
            Cells(4, 17).Value = Greatest_volume
            Cells(4, 16).Value = volume_ticker
            
            
        Next c
        

Next k

                        



        
        
    



    


End Sub



