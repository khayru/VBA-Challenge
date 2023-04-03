Attribute VB_Name = "Module1"
Sub AllYearStock()


    For Each ws In Worksheets
    
        Dim SheetName As String
        'begining  row
        Dim i As Long
        
        'first row of ticker
        Dim j As Long
        
        'count fro fill Ticker row
        Dim TickCounter As Long
        
        'Last row column A
        Dim LastRowA As Long
        
        'last row column I
        Dim LastRowI As Long
        
        'Variable for percent change
        Dim ChangePerc As Double
        
        'Variable for greatest increase
        Dim GIncrease As Double
        
        'Variable for greatest decrease
        Dim GDecrease As Double
        
        'Variable for  total volume
        Dim GVolume As Double
        
        'find work SheetName
        SheetName = ws.Name
        
        'Create headers
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Set Ticker Counter
        TickCounter = 2
        
        
        j = 2
        
        'Find the last cell empty in column A
        
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
            'Loop through all rows
            For i = 2 To LastRowA
            
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I
                ws.Cells(TickCounter, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change in column J
                
                ws.Cells(TickCounter, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating
                    If ws.Cells(TickCounter, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(TickCounter, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(TickCounter, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K
                    If ws.Cells(j, 3).Value <> 0 Then
                    ChangePerc = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(TickCounter, 11).Value = Format(ChangePerc, "Percent")
                    
                    Else
                    
                    ws.Cells(TickCounter, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate total volume in column L
                ws.Cells(TickCounter, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCounter by 1
                TickCounter = TickCounter + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find the last cell empty in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        
        'Prepare for summary
        GVolume = ws.Cells(2, 12).Value
        GIncrease = ws.Cells(2, 11).Value
        GDecrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For greatest total volue if the value is grater then print the value
                
                If ws.Cells(i, 12).Value > GVolume Then
                GVolume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GVolume = GVolume
                
                End If
                
                'For greatest increase if next value is larger print the value
                If ws.Cells(i, 11).Value > GIncrease Then
                GIncrease = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GIncrease = GIncrease
                
                End If
                
                'For greatest decrease  if next value is smaller then print the value
                If ws.Cells(i, 11).Value < GDecrease Then
                GDecrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                GDecrease = GDecrease
                
                End If
                
            'print  summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(GIncrease, "Percent")
            ws.Cells(3, 17).Value = Format(GDecrease, "Percent")
            ws.Cells(4, 17).Value = Format(GVolume, "Scientific")
            
            Next i
            
            
    Next ws
        
End Sub


