Attribute VB_Name = "Module1"
Sub workbook_loop()

Dim i As Integer
Dim ws_num As Integer

Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
ws_num = ThisWorkbook.Worksheets.Count

For i = 1 To ws_num
    ThisWorkbook.Worksheets(i).Activate
    Call StockTracker
    Call StockChampion
Next

starting_ws.Activate 'activate the worksheet that was originally active

End Sub

Sub StockTracker()

  ' Set an initial variable for holding the Stock name
  Dim Stock_Name As String
   Dim Break As String
  Dim StockStart As Double, StockEnd As Double
  ' Set an initial variable for holding the total per credit card Stock
  Dim Stock_Total As Double, lastRow As Double
  Stock_Total = 0
  
  Dim yearlyChange As Double
  
  ' Keep track of the location for each credit card Stock in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  StockStart = Cells(2, 3).Value
  lastRow = Range("a" & Rows.Count).End(xlUp).Row
  ' Set the headers
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  Range("P1").Value = "Ticker"
  Range("Q1").Value = "Volume"
  
  Range("A1:Q1").Font.Bold = True
  Range("A1:Q1").Columns.AutoFit
  Range("k2:K" & lastRow).NumberFormat = "0.00%"
  Range("J2:J" & lastRow).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
  Range("L2:L" & lastRow).NumberFormat = "0.00"
  ' Loop through all credit card purchases
  
  For i = 2 To lastRow
    ' Check if we are still within the same credit card Stock, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Stock name
      Stock_Name = Cells(i, 1).Value
      
      StockEnd = Cells(i, 6).Value
      
      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value

      ' Print the Credit Card Stock in the Summary Table
      Range("I" & Summary_Table_Row).Value = Stock_Name

      yearlyChange = StockEnd - StockStart
      ' Print the Stock Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = yearlyChange
      
      If yearlyChange < 0 Then
        Range("J" & Summary_Table_Row).Interior.Color = vbRed
      Else
        Range("J" & Summary_Table_Row).Interior.Color = vbGreen
      End If
      
        On Error Resume Next
          Range("K" & Summary_Table_Row).Value = (StockEnd - StockStart) / StockStart
        If Err <> 0 Then
            'to infinity and beyond
            Range("K" & Summary_Table_Row).Value = 0
            Err.Clear
        End If
      
      Range("L" & Summary_Table_Row).Value = Stock_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Total
      Stock_Total = 0
        
      StockStart = Cells(i + 1, 3).Value
        
    ' If the cell immediately following a row is the same Stock...
    Else

      ' Add to the Stock Total
      Stock_Total = Stock_Total + Cells(i, 7).Value

    End If

  Next i
  Range("G2:G" & lastRow).Columns.AutoFit
  Range("B2:B" & lastRow).Columns.AutoFit
  
End Sub



Sub StockChampion()

  ' Set an initial variable for holding the Stock name
  Dim Stock_Name As String
  
  Dim IncreaseTicker As String, DecreaseTicker As String, VolumeTicker As String
  Dim IncreasePercent As Double, DecreasePercent As String, VolumeAmount
  
  Dim lastRow As Double
  Stock_Total = 0
  
  
  ' Keep track of the location for each credit card Stock in the summary table
  Dim Summary_Table_Row As Double
  Summary_Table_Row = 2
  
  IncreaseTicker = Cells(2, 9).Value
  IncreasePercent = Cells(2, 10).Value
  
  DecreaseTicker = Cells(2, 9).Value
  DecreasePercent = Cells(2, 10).Value
  

  
  lastRow = Range("I" & Rows.Count).End(xlUp).Row
  'Format the cells
  Range("Q4").NumberFormat = "0.00"
  Range("q2:q3").NumberFormat = "0.00%"
  For i = 2 To lastRow

    ' Check if we are still within the same credit card Stock, if it is not...
    If IncreasePercent < Cells(i, 11).Value Then
        IncreaseTicker = Cells(i, 9).Value
        IncreasePercent = Cells(i, 11).Value
    End If
  
    If VolumeAmount < Cells(i, 12).Value Then
        VolumeTicker = Cells(i, 9).Value
        VolumeAmount = Cells(i, 12).Value
    End If
    
    If DecreasePercent > Cells(i, 11).Value Then
        DecreaseTicker = Cells(i, 9).Value
        DecreasePercent = Cells(i, 11).Value
    End If
    
    
  Next i
    'Format the fake pivot table
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("O2:O4").Columns.AutoFit
    
    Range("O2").Font.Bold = True
    Range("O3").Font.Bold = True
    Range("O4").Font.Bold = True
    
    Range("O1").Columns.AutoFit
    Range("Q2:Q" & lastRow).Columns.AutoFit
    
    
    Range("P2").Value = IncreaseTicker
    Range("P3").Value = DecreaseTicker
    Range("P4").Value = VolumeTicker
    
    Range("Q2").Value = IncreasePercent
    Range("Q3").Value = DecreasePercent
    Range("Q4").Value = VolumeAmount
    
    
End Sub



