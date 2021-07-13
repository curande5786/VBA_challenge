Attribute VB_Name = "Module1"
Sub stock_data()

    'Define variable for worksheet for loop

    Dim ws As Worksheet
    
    'Run worksheet level for loop

    For Each ws In ThisWorkbook.Worksheets
    
        'Activate code on each worksheet

        ws.Activate
        
        'Define variables for column for loop

        Dim Symbol As String

        Dim Volume As Double

        Dim summaryRow As Double
        summaryRow = 2

        Dim yearOpen As Double
        yearOpen = 0

        Dim yearClose As Double
        yearClose = 0

        Dim cellColor As Integer
        
        'Set value of variable to the number of rows on the table
        'with a function that automatically determines the number of rows

        totalRows = Cells(1, 1).End(xlDown).Row
        
        'Filling titles of table headers
    
        ws.Cells(1, 10).Value = "Ticker"
    
        ws.Cells(1, 11).Value = "Yearly Change"
    
        ws.Cells(1, 12).Value = "Percent Change"
    
        ws.Cells(1, 13).Value = "Total Stock Volume"
    
        ws.Cells(2, 16).Value = "Greatest % Increase"
    
        ws.Cells(3, 16).Value = "Greatest % Decrease"
    
        ws.Cells(4, 16).Value = "Greatest Total Volume"
    
        ws.Cells(1, 17).Value = "Ticker"
    
        ws.Cells(1, 18).Value = "Value"
        
        'Run column level for loop

        For r = 2 To totalRows
        
            'Set value of variable to the sum
            'total of the values in a column

            Volume = Volume + Cells(r, 7)
            
            'Run if/then conditional which sets variable to the
            'value of the third column in the current row if this
            'is the first instance of a value in column 1

            If Cells(r - 1, 1).Value <> Cells(r, 1).Value Then
    
                yearOpen = Cells(r, 3)
        
            End If
            
            'Run if/then conditional if this is the
            'last instance of a value in column 1

            If Cells(r + 1, 1).Value <> Cells(r, 1).Value Then
            
                'Set values for variables based on current row
        
                Symbol = Cells(r, 1).Value
                
                yearClose = Cells(r, 6).Value
                
                'Set values of columns to build Summary Table
        
                Range("J" & summaryRow).Value = Symbol
        
                Range("K" & summaryRow).Value = (yearClose - yearOpen)
        
                Range("K" & summaryRow).NumberFormat = "0.00"
                
                'Run if/then conditional which prevents a divide by zero error
                'and fills another column in Summary Table
            
                If yearOpen <> 0 Then
        
                    Range("L" & summaryRow).Value = ((yearClose - yearOpen) / yearOpen)
        
                    Range("L" & summaryRow).NumberFormat = "0.00%"
                
                Else
            
                    Range("L" & summaryRow).Value = 0
        
                End If
                
                'Fill final column of summary table
        
                Range("M" & summaryRow).Value = Volume
        
                summaryRow = summaryRow + 1
        
                Volume = 0
        
            End If
            
            'Run if/then conditional which fills
            'cell colors in a column based on the
            'cell value's relationship with zero

            If Cells(r, 11).Value > 0 Then
    
                cellColor = 4
        
            ElseIf Cells(r, 11).Value < 0 Then
    
                cellColor = 3
        
            ElseIf Cells(r, 11).Value = "" Then
    
                cellColor = 0
        
            End If
        
            Cells(r, 11).Interior.ColorIndex = cellColor
            
            'Run if/then conditional which fills
            'cell colors in a column based on the
            'cell value's relationship with zero
        
            If Cells(r, 12).Value > 0 Then
    
                cellColor = 4
        
            ElseIf Cells(r, 12).Value < 0 Then
    
                cellColor = 3
        
            ElseIf Cells(r, 12).Value = "" Then
    
                cellColor = 0
            
            ElseIf Cells(r, 12).Value = 0 Then
        
                cellColor = 0
        
            End If
    
            Cells(r, 12).Interior.ColorIndex = cellColor
    
        Next r
        
        'Define variables for Greatest Summary table
        
        Dim greatVol As Double
        greatVol = 0
        
        Dim greatUp As Double
        greatUp = 0
        
        Dim greatDown As Double
        greatDown = 0
        
        Dim stock1 As String
        
        Dim stock2 As String
        
        Dim stock3 As String
        
        'Define total rows variable
        
        sumRows = Cells(1, 10).End(xlDown).Row
        
        'Run summary table for loop
        
        For sr = 2 To sumRows
            
            'Run if/then conditional that
            'finds the greatest total volume
            'and corresponding symbol
            
            If Cells(sr, 13).Value > greatVol Then
            
                greatVol = Cells(sr, 13).Value
                
                stock1 = Cells(sr, 10).Value
            
            End If
            
            'Run if/then conditional that
            'finds greatest percent decrease
            'and corresponding symbol
            
            If Cells(sr, 12).Value < greatDown Then
            
                greatDown = Cells(sr, 12).Value
                
                stock2 = Cells(sr, 10).Value
        
            End If
            
            'Run if/then conditional that
            'finds greatest percent increase
            'and corresponding symbol
            
            If Cells(sr, 12).Value >= greatUp Then
            
                greatUp = Cells(sr, 12).Value
                
                stock3 = Cells(sr, 10).Value
                
            End If
            
        Next sr
        
        'Post values found to proper cell
        
        Cells(2, 18).Value = greatUp
        
        Cells(2, 17).Value = stock3
        
        Cells(3, 18).Value = greatDown
        
        Cells(3, 17).Value = stock2
        
        Cells(4, 18).Value = greatVol
        
        Cells(4, 17).Value = stock1
        
        Cells(2, 18).NumberFormat = "0.00%"
        
        Cells(3, 18).NumberFormat = "0.00%"
        
    Next ws

End Sub

