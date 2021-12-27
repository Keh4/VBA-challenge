Sub StockPriceInformation()

    'Set variable for Worksheet
    Dim ws As Worksheet
    
    'Set Headers
    Dim headers()
    
    'Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    headers() = Array("<ticker>", "<date>", "<open>", "<high>", "<low>", "<close>", _
    "<vol>", " ", "Ticker", "Yearly Change", "Percent Change", "Total Stock Value")

    ' Insert header row from first sheet
    Range("A1").Resize(1, UBound(headers) + 1) = headers
    
    'Establish end row for worksheets
        
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set variable for ticker name
  Dim Ticker As String

  ' Set variable for total per ticker
  Dim Day_Volume As Double
  Day_Volume = 0
  
   'Set first variable for Open_Price
  Dim Open_Price As Double
  Open_Price = Cells(2, 3).Value
  
  'Set variable for Close_Price
  Dim Close_Price As Double
  Close_Price = 0

' Set up Summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2


'Set up last row indicator

lastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all ticker daily volumes
  For i = 2 To lastRow

    ' Set Up each loop to only count same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Identify ticker
      Ticker = Cells(i, 1).Value

      ' Add to get Volume Total
      Day_Volume = Day_Volume + Cells(i, 7).Value

        'Find Open Price
         'See above for first Open_Price set at Cell(2,3).Value then Reset below at Open_Price = Cells(i + 1, 3).Value
        
        'Find Close Price
        Close_Price = Cells(i, 6).Value
        
        
      ' Put Ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Put Volume Total to the Summary Table
      Range("L" & Summary_Table_Row).Value = Day_Volume
      
                    'Find Year Change and Put on Summary Table
      Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
      
      'Find Percent Change and Put on Summary Table - Handle the PLNT Error
      
      On Error Resume Next
      
     Range("K" & Summary_Table_Row).Value = (Close_Price - Open_Price) / Open_Price
     
     Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset Volume Total for Next Ticker
      Day_Volume = 0
      
      ' Reset Open_Price for Next Ticker
       Open_Price = Cells(i + 1, 3).Value
      
      ' Reset Close_Price for Next Ticker
      Close_Price = 0

    Else

      ' Add to the Volume Total
      Day_Volume = Day_Volume + Cells(i, 7).Value

    End If

  Next i
  
      'Set variable for Yearly_Change
    Dim Yearly_Change As Double
    
      ' Loop through all Yearly_Change Values
  For i = 2 To lastRow

    'Identify Yearly_Change
    Yearly_Change = Cells(i, 10)
    
    If Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    
    ElseIf Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    
    End If

Next i

  Next ws

End Sub


