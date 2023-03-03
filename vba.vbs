Sub ticker_summary()
Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As String
Dim Total_Stock_Volume As Double
Dim Summary_Row As Integer
Dim Opening As Double
Dim Closing As Double
Dim ws As Worksheet


For Each ws In ThisWorkbook.Sheets
    ws.Select
     
    

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

Total_Stock_Volume = 0

Summary_Row = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).row
summary_lastrow = Cells(Rows.Count, 9).End(xlUp).row

Yearly_Change = 0
Range("J:J").Interior.ColorIndex = 0

    For i = 2 To lastrow
        

        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            Opening = Cells(i, 3).Value
        
        End If
            
            
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Range("I" & Summary_Row).Value = Ticker
        
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            Range("L" & Summary_Row).Value = Total_Stock_Volume
            Total_Stock_Volume = 0
       
            Closing = Cells(i, 6).Value
        
            Yearly_Change = Closing - Opening
            Range("J" & Summary_Row).Value = Yearly_Change
           
            Percent_Change = FormatPercent((Closing / Opening - 1), vbTrue)
            Range("K" & Summary_Row).Value = Percent_Change
        
        If Yearly_Change < 0 Then
            Range("J" & Summary_Row).Interior.Color = vbRed
        Else
            Range("J" & Summary_Row).Interior.Color = vbGreen
        End If
          
            Summary_Row = Summary_Row + 1
            Yearly_Change = 0
            Percent_Change = 0
            Range("J" & Summary_Row).Interior.ColorIndex = 0
        
        Else
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
            
        End If
    
Next i
Dim max_Percent As Double
Dim min_Percent As Double
Dim max_Volume As Double
Dim total_Volume As String
Dim increase As String
Dim decrease As String
Dim Ticker_Increase As String
Dim Ticker_Decrease As String
Dim Ticker_Value As String
Dim row As Integer
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"
Range("P1").Value = "Ticker"
Range("Q1").Value = "Value"

         
max_Percent = WorksheetFunction.Max(Range("K:K"))
increase = FormatPercent(max_Percent, vbTrue)
Range("Q2").Value = increase

min_Percent = WorksheetFunction.Min(Range("K:K"))
decrease = FormatPercent(min_Percent, vbTrue)
Range("Q3").Value = decrease

max_Volume = WorksheetFunction.Max(Range("L:L"))
total_Volume = FormatNumber(max_Volume, vbTrue)
Range("Q4").Value = total_Volume

row = 0

For i = 2 To summary_lastrow
    If Cells(i, 11) = max_Percent Then
    row = i
    Range("P2").Value = Cells(row, 9)
    End If
    If Cells(i, 11) = min_Percent Then
    row = i
    Range("P3").Value = Cells(row, 9)
    End If
    If Cells(i, 12) = max_Volume Then
    row = i
    Range("P4").Value = Cells(row, 9)
    End If
Next i
Next ws
    
End Sub



