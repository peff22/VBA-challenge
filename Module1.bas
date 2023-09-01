Attribute VB_Name = "Module1"
Sub alphabet_practice_real2()

For Each Worksheet In ThisWorkbook.Sheets
        Worksheet.Activate

'   Set initial text variables
Dim Ticker As String


'   Set initial numerical variables for holding totals
Dim Yearly_Change As Double, Percent_Change As Double
Dim open_price As Double, close_price As Double
    

Dim volume As Double
    volume = 0

'   Set up the summary table
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

Dim lastrow As Double
    lastrow = Cells(Rows.Count, "A").End(xlUp).Row
    
Dim i As Long ' new
i = 2 ' new

'Dim Ticker_Title, Yearly, Percent, Total As String
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Set up the Loop
    For i = 2 To (lastrow + 1) 'trying changing this
        
        'Check if we are still within the same ticker symbol
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Set the next ticker symbol
            Ticker = Cells(i, 1).Value
            
            close_price = Cells(i - 1, 6).Value
            open_price = Cells(i, 3).Value ' not sure this is right
            
            'Calculate and assign yearly change
            Yearly_Change = close_price - open_price
            Cells(Summary_Table_Row, 10).Value = Yearly_Change
    
            'Calculate and assign percent change
            If open_price <> 0 Then
                Percent_Change = Yearly_Change / open_price
            Else
                Percent_Change = 0
             
            End If
            
            Cells(Summary_Table_Row, 11).Value = Percent_Change
            
            'Format as percentage
            Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                        
            'Add to the ticker volume total
            volume = volume + Cells(i, 7).Value
            
            'Print the ticker name and volume in the summary table
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("L" & Summary_Table_Row).Value = volume
            
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'reset the volume total
            volume = 0

        Else
        
            volume = volume + Cells(i, 7).Value
        
        End If
        
    Next i
    
    Dim Greatest_Total_Volume As Double: Greatest_Total_Volume = 0
    Dim Greatest_Percent_Increase As Double: Greatest_Percent_Increase = 0
    Dim Greatest_Percent_Decrease As Double: Greatest_Percent_Decrease = 0
    
    lastrow = Cells(Rows.Count, "L").End(xlUp).Row
    For i = 2 To lastrow
        
        Dim current_volume As Double
        current_volume = Cells(i, 12).Value
        
        If current_volume > Greatest_Total_Volume Then
            Greatest_Total_Volume = current_volume
            Range("P4").Value = Cells(i, 9).Value
            
        End If
        
        If Cells(i, 11).Value > Greatest_Percent_Increase Then
            Greatest_Percent_Increase = Cells(i, 11).Value
            ' update the value
            Range("Q2").Value = Cells(i, 11).Value
            ' update the ticker
            Range("P2").Value = Cells(i, 9).Value
        End If
        
        If Cells(i, 11).Value < Greatest_Percent_Decrease Then
            Greatest_Percent_Decrease = Cells(i, 11).Value
            ' update the value
            Range("Q3").Value = Cells(i, 11).Value
            ' update the ticker
            Range("P3").Value = Cells(i, 9).Value
        End If
    Next
    ' update the value
    Range("Q4").Value = Greatest_Total_Volume

    For i = 2 To lastrow
        If Cells(i, 10).Value >= 0.01 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0.01 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    Columns("A:Q").AutoFit
        ' trying without this since I addressed the percentage thing above -- Columns("K:K").NumberFormat = "0.00%"
    
    Next Worksheet

End Sub



