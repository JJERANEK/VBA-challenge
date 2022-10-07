Sub Stocks()

    Dim Ticker As String
    Dim Total_Stock_Volume As Double
    Dim New_Table_Row As String
    Dim Row As Double
    Dim lRow As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Yearly_Change As Double
    Dim WS As Worksheet
         
   For Each WS In Sheets
    WS.Activate
    New_Table_Row = 2
    Total_Stock_Volume = 0
    lRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    Range("I1, P1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Columns("I:L").AutoFit
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("Q1").Value = "Value"
    Columns("O:Q").AutoFit
   
    
    
'Tickers and Volume
    For Row = 2 To lRow
        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
            Ticker = Cells(Row, 1).Value
            Total_Stock_Volume = Total_Stock_Volume + Cells(Row, 7).Value
            Range("I" & New_Table_Row).Value = Ticker
            Range("L" & New_Table_Row).Value = Total_Stock_Volume
            New_Table_Row = New_Table_Row + 1
            Total_Stock_Volume = 0
        Else
            Total_Stock_Volume = Total_Stock_Volume + Cells(Row, 7).Value
        End If
        
    Next Row
    
    
'Yearly Change
    New_Table_Row = 2
            
    For Row = 2 To lRow
        
        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
            Close_Price = Cells(Row, 6).Value
        ElseIf Cells(Row, 1).Value <> Cells(Row - 1, 1).Value Then
            Open_Price = Cells(Row, 3).Value
        End If
                
        If Open_Price > 0 And Close_Price > 0 Then
            Yearly_Change = Close_Price - Open_Price
            Percent_Increase = Yearly_Change / Open_Price
            Range("J" & New_Table_Row).Value = Yearly_Change
            Range("K" & New_Table_Row).Value = FormatPercent(Percent_Increase)
            Close_Price = 0
            Open_Price = 0
            New_Table_Row = New_Table_Row + 1
                    
        End If
      
    Next Row
   
    Max_Percent = WorksheetFunction.Max(ActiveSheet.Columns("K"))
    Min_Percent = WorksheetFunction.Min(ActiveSheet.Columns("K"))
    Max_Volume = WorksheetFunction.Max(ActiveSheet.Columns("L"))
    
    Range("Q2").Value = Max_Percent
    Range("Q2").NumberFormat = "0.00%"
    Range("Q3").Value = Min_Percent
    Range("Q3").NumberFormat = "0.00%"
    Range("Q4").Value = Max_Volume
    
    For Row = 2 To lRow
        If Max_Percent = Cells(Row, "K").Value Then
            Range("P2").Value = Cells(Row, "I").Value
        ElseIf Min_Percent = Cells(Row, "K").Value Then
            Range("P3").Value = Cells(Row, "I").Value
        ElseIf Max_Volume = Cells(Row, "L").Value Then
            Range("P4").Value = Cells(Row, "I").Value
        End If
    Next Row
   
   For Row = 2 To lRow
        If IsEmpty(Cells(Row, "J").Value) Then Exit For
        If Cells(Row, "J").Value > 0 Then
            Cells(Row, "J").Interior.ColorIndex = 4
        Else
            Cells(Row, "J").Interior.ColorIndex = 3
        End If
    Next Row
    
  Next WS

    

End Sub
