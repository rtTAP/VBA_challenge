Attribute VB_Name = "Module1"
Option Explicit
Sub Multi_Year_Stock_Report()

'set a variable for the worksheet
Dim ws As Worksheet
 'WorksheetName = ("A")
 'Set ws = ThisWorkbook.Worksheets
 'Dim Current As Worksheet
 'For Each Current In Worksheets

 
 For Each ws In ActiveWorkbook.Worksheets
 ws.Activate
'For Each ws In ThisWorkbook.Worksheets
 'For Each ws In ActiveWorkbook.Worksheets

'Print the name of each required column
 Range("I1").Value = "Ticker"
 Range("J1").Value = "Yearly Change"
 Range("K1").Value = "Percent Change"
 Range("L1").Value = "Total Stock Volume"
 
 
'set an initial variable for the ticker
  Dim Ticker_Name As String
 

'set variable Ticker Total Volume
  Dim Ticker_TotalVolume As Double
  Ticker_TotalVolume = 0

'set a variable for Yearly Change
  Dim Yearly_Change As Double
  Yearly_Change = 0

'set a variable for Ticker open Initial open value
  Dim Ticker_OpenValue As Double
  'Ticker_OpenValue = 0 _ NOT 0
  Ticker_OpenValue = Cells(2, 3).Value
 
 'set a variable for Ticker close
  Dim Ticker_CloseValue As Double
  'Ticker_CloseValue = 0
 
 
'set a variable for percent change
  Dim Percent_Change As Double
  Percent_Change = 0


'keep track of each ticker in summary table
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2

'loop through all tickers
 Dim i As Long
 'For i = 2 To 759001
 For i = 2 To Range("A2").End(xlDown).Row
 
   'Check if values are still same Ticker Name
   If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   

    'Set the Ticker Name
    Ticker_Name = Cells(i, 1).Value
    
    'Calculate Ticker Total Volume
    Ticker_TotalVolume = Ticker_TotalVolume + Cells(i, 7)
    
    'Find Ticker Close Value Initial Close Value
    'Ticker_CloseValue = Ticker_CloseValue + Cells(i, 6).Value
     Ticker_CloseValue = Cells(i, 6).Value
   
     
     'Yearly Change = OpenValue - CloseValue
      Yearly_Change = Ticker_CloseValue - Ticker_OpenValue
    
     
     'Calculate Ticker Percentage Change
     Percent_Change = Yearly_Change / Ticker_OpenValue
    
      
     'Find Ticker Open Value
     Ticker_OpenValue = Cells(i + 1, 3).Value
     
    'Print Ticker Name in Summary Table
     Range("I" & Summary_Table_Row).Value = Ticker_Name
    
    'Print the Yearly Change
     Range("J" & Summary_Table_Row).Value = Yearly_Change
     
     'Print Percentage Change
      Range("K" & Summary_Table_Row).Value = Percent_Change
     
     'Print Ticker Open Value
     'Range("O" & Summary_Table_Row).Value = Ticker_OpenValue
    
    'Print the Ticker Close Value
     'Range("P" & Summary_Table_Row).Value = Ticker_CloseValue
    
    'Print the Ticker Total Volume in Summary Table
      Range("L" & Summary_Table_Row).Value = Ticker_TotalVolume
      
     'Add one row to the summary table
     Summary_Table_Row = Summary_Table_Row + 1
     
     'Reset the Ticket Total Volume
     Ticker_TotalVolume = 0
     
    
     
 'if cell immediately following that row is the same Ticker Name then do something else
 Else
 
  'add to the Ticker Total Volume
  Ticker_TotalVolume = Ticker_TotalVolume + Cells(i, 7).Value
  
  'add to the Ticker Open Value
  'Ticker_OpenValue = Ticker_OpenValue + Cells(i, 3).Value
  
  'add to the Ticker Close Value
  'Ticker_CloseValue = Ticker_CloseValue + Cells(i, 6).Value
  
  'Yearly_Change = Ticker_CloseValue - Ticker_OpenValue
  
    
    End If


 Next i
  
 
'ws.Activate
Debug.Print ws.Name
Next


End Sub




