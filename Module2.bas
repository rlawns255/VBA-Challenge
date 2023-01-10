Attribute VB_Name = "Module1"
Sub stocks()
    'to loop through all sheets
    For Each ws In Worksheets
        ' Setting an initial variable for ticker
        Dim Ticker_name As String
    
        'Setting an initial variable for Ticker_total
        Dim Ticker_total As Double
        Ticker_total = 0
    
        'Setting an Initial Variable for Yearly Change
        Dim Yearly_change As Double
        Initial_Yearly_change = ws.Cells(2, 3).Value
    
        'Setting an Initial Variable for Percent change
        Dim Percent_change As Double
    
        'Location tracker for Ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        'setting the summary table header
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        'determining the last row
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Loop through all Stock changes
        For i = 2 To LastRow
    
            'Checking if same ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
                'setting Ticker name
                Ticker_name = ws.Cells(i, 1).Value
            
                'Print the Ticker name in the summary table
            
                ws.Range("I" & Summary_Table_Row).Value = Ticker_name
            
                'adding Ticker total
                Ticker_total = Ticker_total + ws.Cells(i, 7).Value
            
                'Print the Stock Total in the summary table
            
                ws.Range("L" & Summary_Table_Row).Value = Ticker_total
            
                'calculating Yearly change
                Yearly_change = ws.Cells(i, 6).Value - Initial_Yearly_change
            
                'calculating percent change
                Percent_change = (Yearly_change / Initial_Yearly_change)
            
                'Print the Yearly change in the summary table
                ws.Range("J" & Summary_Table_Row).Value = Yearly_change
            
                'Print the Percent change in the summary table
                ws.Range("K" & Summary_Table_Row).Value = Percent_change
            
                'Reinitiating Initial Yearly Change
                Initial_Yearly_change = ws.Cells(i + 1, 6).Value
            
                'add oine to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
            
                'reset all variable
                Ticker_total = 0
            
                Yearly_change = 0
            
                Percent_change = 0
            
            Else
                'add to the ticker total
                Ticker_total = Ticker_total + ws.Cells(i, 7).Value
        
            End If
        
        Next i
    
        'setting the Fuctionality Table
        ws.Range("O" & 2) = "Greatest % Increase"
        ws.Range("O" & 3) = "Greatest % Decrease"
        ws.Range("O" & 4) = "Greatest Total Volume"
        ws.Range("P" & 1) = "Ticker"
        ws.Range("Q" & 1) = "Value"
    
        'setting Functionality table row
        Dim Function_Table_Row As Integer
        Function_Table_Row = 2
        
        'getting last row of Summary Table
        Dim SLastRow As Long
        SLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'variables
        Dim GreatestPercent As Double
        Dim LowestPercent As Double
        Dim GreatestStock As Double
        
        GreatestPercent = ws.Cells(2, 11).Value
        LowestPercent = ws.Cells(2, 11).Value
        GreatestStock = ws.Cells(2, 12).Value
        
        'looping through Summary
        
        For y = 2 To SLastRow
            'conditions for Greatest percent increase
            If ws.Cells(y, 11).Value >= GreatestPercent Then
            GreatestPercent = ws.Cells(y, 11).Value
            
            'print
            ws.Cells(2, 16).Value = Cells(y, 9).Value
            ws.Cells(2, 17).Value = GreatestPercent
       
            'conditions for lowest percent decrease
            ElseIf ws.Cells(y, 11).Value <= LowestPercent Then
            LowestPercent = ws.Cells(y, 11).Value
            
            'print
            ws.Range("P" & 3).Value = Cells(y, 9).Value
            ws.Range("Q" & 3).Value = LowestPercent
            
            'conditions for greatest stock volume
            ElseIf ws.Cells(y, 12).Value >= GreatestStock Then
            GreatestStock = ws.Cells(y, 12).Value
            
            'print
            ws.Range("P" & 4).Value = Cells(y, 9).Value
            ws.Range("Q" & 4).Value = GreatestStock
            
            End If
        
        Next y
            
        
    Next ws
    
End Sub
