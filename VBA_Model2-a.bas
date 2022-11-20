Attribute VB_Name = "Module1"
Sub vba_homework()


For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'Value for ranges
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly_Change"
Range("K1").Value = "Percentage_Change"
Range("L1").Value = "Volume"



  ' Set an initial variable for ticker
    Dim Ticker As String
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Percetage_Change As Double
    Percentage_Change = 0
    'Dim Opening_Price As Double
    'Opening_Price = 0
    Dim Closing_Price As Double
    Closing_Price = 0
    Dim Volume As Double
    
    Opening_Price = Cells(2, 3).Value
    
'find end of row
lRow = Cells(Rows.Count, 1).End(xlUp).Row

    ' Keep track of the location for each ticker in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    ' Loop through all Tickers
    For i = 2 To lRow

    ' Check if we are still within the same ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        Ticker = Cells(i, 1).Value
        Volume = Volume + Cells(i, 7).Value
 
        Closing_Price = Cells(i, 6).Value
        Yearly_Change = Closing_Price - Opening_Price
        Percentage_Change = (Yearly_Change / Opening_Price) * 100
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("J" & Summary_Table_Row).Value = Yearly_Change
        Range("K" & Summary_Table_Row).Value = Percentage_Change
        Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        Range("L" & Summary_Table_Row).Value = Volume
    
               
    ' Add one to the summary table row
    Opening_Price = Cells(i + 1, 3).Value
    Summary_Table_Row = Summary_Table_Row + 1
      
    ' Reset the volume
    Volume = 0

    ' If the cell immediately following a row is the same brand...
    Else

    ' Add to the Brand Total
      Volume = Volume + Cells(i, 7).Value
      
        If (Yearly_Change >= 0) Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        
        End If

        If (Yearly_Stock_Change_Percent > Greatest_Increase_Ticker_Percent) Then
            Greatest_Increase_Ticker_Percent = Yearly_Stock_Change_Percent
            Greatest_Increase_Ticker = Ticker_Name
                    
        ElseIf (Yearly_Stock_Change_Percent < Greatest_Decrease_Ticker_Percent) Then
            Greatest_Decrease_Ticker_Percent = Yearly_Stock_Change_Percent
            Greatest_Decrease_Ticker = Ticker_Name
                    
        End If
                       
        If (Ticker_Total_Volume > Greatest_Volume) Then
            Greatest_Volume = Ticker_Total_Volume
            Greatest_Volume_Ticker = Ticker_Name
        End If

    End If
    


Next i

    'adjust column width
    Columns("I:Q").AutoFit

Next ws

End Sub

