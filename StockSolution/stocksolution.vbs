Sub ticker_count()

'Set an initial variables
Dim stock As String
Dim TotalChange As Double
Dim Perchange As Double
Dim i As Double
Dim j As Double
Dim volume As Double
Dim ws as Worksheet
Dim Summary_Table_Row As Long

for each ws in ActiveWorkbook.Worksheets

 'declare an initial variables
volume = 0
TotalChange = 0
j = 2

 ' Keep track of the location for each Ticker in the summary table
  Summary_Table_Row = 2
  
  'Calculate the last row
  row_count = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all stocks
  For i = 2 To Str(row_count)
  volume = volume + ws.Cells(i, 7).Value
      
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the Ticker name
      stock = ws.Cells(i, 1).Value

      ' Calculate Total change
      TotalChange = TotalChange + (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value)
      
      If ws.Cells(j, 3).Value = 0 Then
      
      Perchange = 0
      
      Else
      'Percentage change
       Perchange = (TotalChange / ws.Cells(j, 3).Value)
       
       End If
       
      
      j = i + 1
      
      ' Print the Ticker name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = stock

      ' Print the Volume amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = volume

    ' Print Total change in the summary table
    
      ws.Range("J" & Summary_Table_Row).Value = TotalChange
      
     ' Print Perchange in the summary table
     
     ws.Range("K" & Summary_Table_Row).Value = Perchange
     ws.Range("K:K").Style = "Percent"

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset everything
      volume = 0
TotalChange = 0

    
    
End If




  Next i
Next ws
End Sub




Sub colorcode()
Dim j As Double
Dim ws As Worksheet
Dim colour_row_count As Double


For Each ws In ActiveWorkbook.Worksheets

colour_row_count = ws.Cells(Rows.Count, 11).End(xlUp).Row

For j = 2 To Str(colour_row_count)

If ws.Cells(j, 11).Value < 0 Then

ws.Cells(j, 11).Interior.ColorIndex = 3


Else
ws.Cells(j, 11).Interior.ColorIndex = 4
End If

Next j

Next ws

End Sub




Sub largest()
Dim rng, ticker_rng, rng_volume As Range
Dim ws As Worksheet
Dim large As Double
Dim min As Double
Dim ticker_min, ticker_large, ticker_high_vol As String

For Each ws In ActiveWorkbook.Worksheets

Set rng = ws.Range("K2:K9000")
Set ticker_rng = ws.Range("I2:I9000")
Set rng_volume = ws.Range("L2:L9000")

large = Application.WorksheetFunction.Max(rng)
min = Application.WorksheetFunction.min(rng)
high_vol = Application.WorksheetFunction.Max(rng_volume)

  ticker_large = WorksheetFunction.Index(ticker_rng, WorksheetFunction.match(WorksheetFunction.Max(rng), rng, False))
  ticker_min = WorksheetFunction.Index(ticker_rng, WorksheetFunction.match(WorksheetFunction.min(rng), rng, False))
  ticker_high_vol = WorksheetFunction.Index(ticker_rng, WorksheetFunction.match(WorksheetFunction.Max(rng_volume), rng_volume, False))
  

 
 ws.Cells(5, 15).Value = ticker_large
 ws.Cells(5, 16).Value = large
 ws.Cells(5, 16).Style = "Percent"
  ws.Cells(6, 15).Value = ticker_min
 ws.Cells(6, 16).Value = min
 ws.Cells(6, 16).Style = "Percent"
 ws.Cells(7, 15).Value = ticker_high_vol
 ws.Cells(7, 16).Value = high_vol
 
Next ws
 
 
End Sub