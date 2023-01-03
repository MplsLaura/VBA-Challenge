Sub Summarize_Sheets()

    Dim Ticker As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim First_Row As Long
    Dim Last_Row As Long
    Dim Summary_Table_Row As Integer
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Total_Volume As Double
    Dim Greatest_Increase_Ticker As String
    Dim Greatest_Decrease_Ticker As String
    Dim Greatest_Total_Volume_Ticker As String
    Dim Significant_Values_Table_Row As Double
        
    Worksheet_Count = ActiveWorkbook.Worksheets.Count
    For Sheet = 1 To Worksheet_Count
 
        Worksheets(Sheet).Activate
        Yearly_Change = 0
        Percent_Change = 0
        Total_Stock_Volume = 0
        First_Row = 2
        Last_Row = Cells(Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
    
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Total_Volume = 0
        Significant_Values_Table_Row = 1
        
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Stock Volume"
    
        For i = 2 To Last_Row
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
                Ticker = Cells(i, 1).Value
                
                Yearly_Change = (Cells(i, 6).Value - Cells(First_Row, 3).Value)
                
                Percent_Change = (Cells(i, 6).Value - Cells(First_Row, 3).Value) / Cells(First_Row, 3).Value
                
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
                Cells(Summary_Table_Row, 9).Value = Ticker
                Cells(Summary_Table_Row, 10).Value = Yearly_Change
                
                If (Yearly_Change >= 0) Then
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                Else
                    Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
                    
                End If
                
                Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
                Cells(Summary_Table_Row, 11).Value = Percent_Change
    
                Cells(Summary_Table_Row, 12).Value = Total_Stock_Volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                
    
                
                If (Greatest_Increase < Percent_Change) Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = Ticker
                    
                End If
                
                If (Greatest_Decrease > Percent_Change) Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = Ticker
                    
                End If
                
                If (Greatest_Total_Volume < Total_Stock_Volume) Then
                    Greatest_Total_Volume = Total_Stock_Volume
                    Greatest_Total_Volume_Ticker = Ticker
                End If
                
                Total_Stock_Volume = 0
                First_Row = i + 1
                
            Else
            
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
            End If
            
    
        Next i
                
        Cells(Significant_Values_Table_Row, 16).Value = "Ticker"
        Cells(Significant_Values_Table_Row, 17).Value = "Value"
        
        Significant_Values_Table_Row = Significant_Values_Table_Row + 1
        
        Cells(Significant_Values_Table_Row, 15).Value = "Greatest % Increase"
        Cells(Significant_Values_Table_Row, 16).Value = Greatest_Increase_Ticker
        Cells(Significant_Values_Table_Row, 16).NumberFormat = "0.00%"
        Cells(Significant_Values_Table_Row, 17).Value = Greatest_Increase
        
        Significant_Values_Table_Row = Significant_Values_Table_Row + 1
        
        Cells(Significant_Values_Table_Row, 15).Value = "Greatest % Decrease"
        Cells(Significant_Values_Table_Row, 16).Value = Greatest_Decrease_Ticker
        Cells(Significant_Values_Table_Row, 16).NumberFormat = "0.00%"
        Cells(Significant_Values_Table_Row, 17).Value = Greatest_Decrease
        
        Significant_Values_Table_Row = Significant_Values_Table_Row + 1
        
        Cells(Significant_Values_Table_Row, 15).Value = "Greatest Total Volume"
        Cells(Significant_Values_Table_Row, 16).Value = Greatest_Total_Volume_Ticker
        Cells(Significant_Values_Table_Row, 17).Value = Greatest_Total_Volume
        
        Next Sheet


    
End Sub




