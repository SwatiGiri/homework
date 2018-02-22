Sub share()

    Dim ws As Worksheet
    For Each ws In Worksheets
    
           ' Determine the Last Row
      LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
      
      
      ws.Cells(1, 10).Value = "Ticker"
      ws.Cells(1, 11).Value = "Yearly Change"
      ws.Cells(1, 12).Value = "Percentage Change"
      ws.Cells(1, 13).Value = "Total Stock Volume"
      ws.Cells(1, 16).Value = "Ticker"
      ws.Cells(1, 17).Value = "Value"
      ws.Cells(2, 15).Value = "Greatest % Increase"
      ws.Cells(3, 15).Value = "Greatest % Decrease"
      ws.Cells(4, 15).Value = "Greatest Total Volume"
      
      ' Set an initial variable for holding the Share
      Dim Ticker_Name As String
    
      ' Set an initial variable for holding the total Share
      Dim Ticker_Total As Double
      Ticker_Total = 0
      
      'Set an initial variable for holding the yearly share change
      Dim Yearly_Change As Double
      Yearly_Change = 0
      
      'Set an initial variable for holding the percentage share change
      Dim Percentage_Change As Double
      Percentage_Change = 0
      
      ' Keeps track of the greatest % increase
      Dim Greatest_Increase As Double
      Dim Greatest_Increase_Company As String ' Keeps track of the company name
      ' Keeps track of the greatest % Decrease
      Dim Greatest_Decreaset As Double
      Dim Greatest_Decrease_Company As String
      ' Keeps track of the greatest % Decrease
      Dim Greatest_Total_Volume As Double
      Dim Greatest_Total_Volume_Company As String
    
      ' Keep track of the location for each credit card Ticker in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
            
      ' Keeps track of the opening day for every stock
      Dim Opening_Day As Double
      
      
      ' Loop through all shares purchased
      For i = 2 To LastRow
        If i = 2 Then
            Opening_Day = Cells(i, 3).Value
        End If
        
        ' Check if we are still within the same  Ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          ' Set the Share name
          Ticker_Name = Cells(i, 1).Value
    
          ' Add to the Share Total
          Ticker_Total = Ticker_Total + Cells(i, 7).Value
          If Greatest_Total_Volume < Ticker_Total Then
            Greatest_Total_Volume = Ticker_Total
            Greatest_Total_Volume_Company = Cells(i, 1).Value
          End If
          
          'Yearly change from what the stock opened the year at to what the closing price was
          Yearly_Change = Cells(i, 6) - Opening_Day
          
          'The percent change from the what it opened the year at to what it closed.
          If Yearly_Change = 0 Or Opening_Day = 0 Then
            Percentage_Change = 0
          Else
            Percentage_Change = Round((Yearly_Change / Opening_Day) * 100, 2)
          End If
          
          If Percentage_Change > Greatest_Increase Then
            Greatest_Increase = Percentage_Change
            Greatest_Increase_Company = Cells(i, 1).Value
          End If
          
          If Percentage_Change < Greatest_Decrease Then
            Greatest_Decrease = Percentage_Change
            Greatest_Decrease_Company = Cells(i, 1).Value
          End If
          
          
          ' Print the Share in the Summary Table
          Range("J" & Summary_Table_Row).Value = Ticker_Name
    
          ' Print the Ticker Amount to the Summary Table
          Range("M" & Summary_Table_Row).Value = Ticker_Total
          
          'Print the Ticker Yearly Change to the Summary Table
          Range("K" & Summary_Table_Row).Value = Yearly_Change
          'Conditional Formatting, changing bg color
          If Yearly_Change < 0 Then
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
          Else
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
          End If
          
           'Print the Ticker Yearly Change to the Summary Table
          Range("L" & Summary_Table_Row).Value = CStr(Percentage_Change) + "%"
          
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Ticker Total , Ticker Change , Percentage Change
          Ticker_Total = 0
          Yearly_Change = 0
          Percentage_Change = 0
          Opening_Day = Cells(i + 1, 3).Value
          
          
    
        ' If the cell immediately following a row is the same Ticker...
        Else
        
          ' Add to the Ticker Total
          Ticker_Total = Ticker_Total + Cells(i, 7).Value
        
              
        End If
    
      Next i
      
      Cells(2, 17).Value = CStr(Greatest_Increase) + "%"
      Cells(3, 17).Value = CStr(Greatest_Decrease) + "%"
      Cells(4, 17).Value = Greatest_Total_Volume
      Cells(2, 16).Value = Greatest_Increase_Company
      Cells(3, 16).Value = Greatest_Decrease_Company
      Cells(4, 16).Value = Greatest_Total_Volume_Company
    Next

End Sub
