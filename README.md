# VBA---Challenge---PG

Sub stockdata()

    For Each ws In Worksheets
    
        'dim variables
        
        Dim TCOUNT As Long      'variable for ticker counter
        
        Dim LROTA As Long       'vairbale for last row of the column A for the ticker, for the last value of the loop
        
        Dim LROTJ As Long       'variable for last row of the column J for the ticker, for the last value of the loop
        
        Dim PC As Double        'variable for obtaining the value of the percentage change
        
        Dim GI As Double        'variable for obtaining the greatest % increase
        
        Dim GD As Double        'variable for obtaining the greatest % decrease
        
        Dim GV As Double        'variable for obtaining the greatest total volume
        
        Dim i As Long      'variable for getting the actual ticker row
        
        Dim j As Long     'variable to establish the starting row of ticker
        
        
        'Create column headers for each worksheet
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percentage Change"
        ws.Range("M1").Value = "Total Stock Volume"
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
    
        
        'setting ticker counter to first row of relevant data
        TCOUNT = 2
        
        'setting start row to 2
        j = 2
        
        'last row with ticker information to use it in the cycle
        LROTA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
            'Loopping through all the rows of the data base
            For i = 2 To LROTA
                
                'condition for the loop
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Writting ticker in column J
                ws.Cells(TCOUNT, 10).Value = ws.Cells(i, 1).Value
                
                'Calculating Yearly Change in column K
                ws.Cells(TCOUNT, 11).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    
                    'Conditional formating to assign color depending on the criteria
                    If ws.Cells(TCOUNT, 11).Value < 0 Then
                
                    'criteria for color red
                    ws.Cells(TCOUNT, 11).Interior.ColorIndex = 3
                
                    Else
                
                    'criteria for color green
                    ws.Cells(TCOUNT, 11).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculating percent change in column L
                    If ws.Cells(j, 3).Value <> 0 Then
                    PC = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TCOUNT, 12).Value = Format(PC, "Percent")
                    
                    End If
                     
                'Calculating total stock volume in column M
                ws.Cells(TCOUNT, 13).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
            
                'Increasing TCOUNT by 1
                TCOUNT = TCOUNT + 1
                
                'new starting row for j
                j = j + 1
                
                End If
            
            Next i
            
    
        'last row of column J
        LROTJ = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'setting things for analyzing the stats created table
        GV = ws.Cells(2, 13).Value
        GI = ws.Cells(2, 12).Value
        GD = ws.Cells(2, 12).Value
          
          
            'looping through the stats created table
            For i = 2 To LROTJ
            
           
                
                'analyzing GV, and comparing in order to get the largest one in each iteration
                If ws.Cells(i, 13).Value > GV Then
                GV = ws.Cells(i, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
                
                Else
                
                GV = GV
                
                End If
                
                'analyzing GI, and comparing in order to get the largest one in each iteration
                If ws.Cells(i, 12).Value > GI Then
                GI = ws.Cells(i, 12).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
                
                Else
                
                GI = GI
                
                End If
                
                'analyzing GD, and comparing in order to get the largest one in each iteration
                If ws.Cells(i, 12).Value < GD Then
                GD = ws.Cells(i, 12).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
                
                Else
                
                GD = GD
                
                End If
                
            'return stock for GI, GD and GV with the correct format
            ws.Range("R2").Value = Format(GI, "Percent")
            ws.Range("R3").Value = Format(GD, "Percent")
            ws.Range("R4").Value = Format(GV, "Scientific")
            
            Next i
              
    
    'jumping into the next worksheet
    Next ws
        
End Sub
                
        

