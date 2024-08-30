
Option Explicit
Sub Multi_Year_Stock_Report()

    'set a variable for the worksheet
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
        
        'Print the name of each required column
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'set an initial variable for the ticker
        Dim Ticker_Name As String
        
        'set variable Ticker Total Volume
        Dim Ticker_TotalVolume As Double
        Ticker_TotalVolume = 0
        
        'set a variable for Yearly Change
        Dim Yearly_Change As Double
        
        'set a variable for Ticker open Initial open value
        Dim Ticker_OpenValue As Double
        Ticker_OpenValue = ws.Cells(2, 3).Value
        
        'set a variable for Ticker close
        Dim Ticker_CloseValue As Double
        
        'set a variable for percent change
        Dim Percent_Change As Double
        
        'Variables to track the greatest values
        Dim Greatest_Increase As Double
        Dim Greatest_Decrease As Double
        Dim Greatest_Volume As Double
        Dim Greatest_Increase_Ticker As String
        Dim Greatest_Decrease_Ticker As String
        Dim Greatest_Volume_Ticker As String
        
        Greatest_Increase = 0
        Greatest_Decrease = 0
        Greatest_Volume = 0
        
        'keep track of each ticker in summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'loop through all tickers
        Dim i As Long
        For i = 2 To ws.Range("A2").End(xlDown).Row
            
            'Check if values are still the same Ticker Name
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set the Ticker Name
                Ticker_Name = ws.Cells(i, 1).Value
                
                'Calculate Ticker Total Volume
                Ticker_TotalVolume = Ticker_TotalVolume + ws.Cells(i, 7).Value
                
                'Find Ticker Close Value
                Ticker_CloseValue = ws.Cells(i, 6).Value
                
                'Calculate Yearly Change
                Yearly_Change = Ticker_CloseValue - Ticker_OpenValue
                
                'Calculate Ticker Percentage Change
                If Ticker_OpenValue <> 0 Then
                    Percent_Change = Yearly_Change / Ticker_OpenValue
                Else
                    Percent_Change = 0
                End If
                
                'Print Ticker Name in Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                
                'Print the Yearly Change
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                'Print Percentage Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                
                'Print the Ticker Total Volume in Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_TotalVolume
                
                'Check for Greatest Increase/Decrease in Percent Change
                If Percent_Change > Greatest_Increase Then
                    Greatest_Increase = Percent_Change
                    Greatest_Increase_Ticker = Ticker_Name
                End If
                
                If Percent_Change < Greatest_Decrease Then
                    Greatest_Decrease = Percent_Change
                    Greatest_Decrease_Ticker = Ticker_Name
                End If
                
                'Check for Greatest Total Volume
                If Ticker_TotalVolume > Greatest_Volume Then
                    Greatest_Volume = Ticker_TotalVolume
                    Greatest_Volume_Ticker = Ticker_Name
                End If
                
                'Add one row to the summary table
                Summary_Table_Row = Summary_Table_Row + 1
                
                'Reset the Ticker Total Volume
                Ticker_TotalVolume = 0
                
                'Reset the Ticker Open Value for the next ticker
                If ws.Cells(i + 1, 1).Value <> "" Then
                    Ticker_OpenValue = ws.Cells(i + 1, 3).Value
                End If
                
            Else
                'add to the Ticker Total Volume
                Ticker_TotalVolume = Ticker_TotalVolume + ws.Cells(i, 7).Value
                
            End If
            
        Next i
        
       ' Apply Conditional Formatting to the Yearly Change column (J)
        ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Delete
        ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions(1).Interior.Color = vbGreen
        
        ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range("J2:J" & Summary_Table_Row - 1).FormatConditions(2).Interior.Color = vbRed
        
        ' Apply Conditional Formatting to the Percent Change column (K)
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions.Delete
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions(1).Interior.Color = vbGreen
        
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        ws.Range("K2:K" & Summary_Table_Row - 1).FormatConditions(2).Interior.Color = vbRed
        
        ' Print the greatest values
        ws.Range("P2").Value = "Greatest % Increase"
        ws.Range("P3").Value = "Greatest % Decrease"
        ws.Range("P4").Value = "Greatest Total Volume"
        
        ws.Range("Q2").Value = Greatest_Increase_Ticker
        ws.Range("R2").Value = Greatest_Increase
        ' Format as percentage
        ws.Range("R2").NumberFormat = "0.00%"
        
        ws.Range("Q3").Value = Greatest_Decrease_Ticker
        ws.Range("R3").Value = Greatest_Decrease
        ' Format as percentage
        ws.Range("R3").NumberFormat = "0.00%"
        
        ws.Range("Q4").Value = Greatest_Volume_Ticker
        ws.Range("R4").Value = Greatest_Volume
        
    Next ws

End Sub




