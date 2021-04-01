Attribute VB_Name = "Module11"
Sub Drew_stock_market_analysis()

'loop through worksheets

For Each ws In Worksheets

    'declare variables
    
    Dim current_row As Long
    Dim Last_row As Long
    Dim summary_row As Integer
    Dim Total_Volume
    Dim Symbol As String
    Dim Opening As Double
    
    ws.Activate
    Range("I:L").Clear
    
    'initialize variables
    Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    summary_row = 2
    Total_Volume = 0
    
    'Create headers and summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Columns("I:L").Select
   ' Selection.ColumnWidth = 13
    'ws.Range("M15").Select
    

    ' Iterate through the worksheet from Row 2 to last row
        For current_row = 2 To Last_row
        If Symbol <> ws.Cells(current_row, 1) Then
            Symbol = ws.Cells(current_row, 1)
            Opening = ws.Cells(current_row, 3)
        End If
            
       Total_Volume = Total_Volume + ws.Cells(current_row, 7).Value
            
        If ws.Cells(current_row + 1, 1) <> ws.Cells(current_row, 1) Then
            ws.Range("L" & summary_row).Value = Total_Volume
            
   
         
            'print ticket to summary table
            ws.Range("I" & summary_row).Value = ws.Cells(current_row, 1).Value
            
    ' Calculate yearly change
       Cells(summary_row, 10) = Cells(current_row, 6) - Opening
             
     ' Calculate Percent Change
     If Opening = 0 Then
     Cells(summary_row, 11) = 0
     Else
        Cells(summary_row, 11) = Cells(summary_row, 10) / Opening
     End If
     Cells(summary_row, 11).NumberFormat = "0.00%"
     
        'Color coding check
        
        If Cells(summary_row, 11) < 0 Then
            Cells(summary_row, 11).Interior.ColorIndex = 3
            Else
            Cells(summary_row, 11).Interior.ColorIndex = 4
        End If
        
        
            Total_Volume = 0
            summary_row = summary_row + 1
            
            
            
        End If
            
    
    Next current_row


Next ws


MsgBox "Complete"

End Sub

