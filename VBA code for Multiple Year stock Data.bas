Attribute VB_Name = "Module1"
Sub Test_current_Macro()

    'Declare Variables
     Dim ws As Worksheet
     Dim Ticker As String
     Dim Volume As Double
     Dim Change As Double
     Dim Percent As Double
     
     
     'Begin Iterattion through each worksheet
     For Each ws In ThisWorkbook.Worksheets
     
     
     

    'create summary table headers
     ws.Range("I1").Value = "Ticker"
     ws.Range("J1").Value = "Yearly Change"
     ws.Range("K1").Value = "Percent Change"
     ws.Range("L1").Value = "Total Stock Volume"
     ws.Range("P1").Value = "Ticker"
     ws.Range("Q1").Value = "Value"
     ws.Range("O2").Value = "Greatest % Increase"
     ws.Range("O3").Value = "Greatest % Decrease"
     ws.Range("O4").Value = "Greatest Total Volume"
     'Formatting
     ws.Columns("I:Q").EntireColumn.AutoFit
     
    'Initialize variables
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    j = 2
    
    Volume = 0
    Start = 2
    Change = 0
    Percent = 0
    
    'Iterrate
        For i = 2 To lastRow
            
            'Check if Ticker Changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Put Previous cell value in Ticker Column
                ws.Range("I" & j).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change
                Change = ws.Cells(i, 6).Value - ws.Cells(Start, 6).Value
                
                'Input Change
                ws.Range("J" & j).Value = Change
                
                'Calculate percent change
                Percent = Change / ws.Cells(Start, 6).Value
                
                'Record Change
                ws.Range("K" & j).Value = Percent
                
                'Style cells
                ws.Range("K" & j).NumberFormat = "0.00%"
                
                'combine volumes
                Volume = Volume + ws.Cells(i, 7).Value
                
                'Place total in Total stock volume
                ws.Range("L" & j).Value = Volume
                
                'Incriment j by 1
                j = j + 1
                
                'Reset Volume
                Volume = 0
                
                'Reset Change
                Change = 0
                
                'Incriment Start
                Start = i
                
            'Add volume for each ticker
            Else
                'combine volumes
                Volume = Volume + ws.Cells(i, 7).Value
                
                
                
            End If
            
            
            
        Next i
    
        
    
    
    'Find greatest % increase

    
    Dim lastRowPer As Long
    Dim maxinc As Double
    Dim incTicker As String
    Dim tickerRange As Range
    Dim maxRange As Range
    

    lastRowPer = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row ' get the last row in column I
    
    Set tickerRange = ws.Range("I2:I" & lastRowPer) ' set the range for column I
    Set maxRange = ws.Range("K2:K" & lastRowPer) ' set the range for column K
    
    maxinc = WorksheetFunction.Max(maxRange) ' get the maximum value from column K
    incTicker = tickerRange.Cells(Application.WorksheetFunction.Match(maxinc, maxRange, 0)).Value ' get the corresponding ticker from column I
    
    ws.Range("P2").Value = incTicker ' set the value in P2 to the max ticker
    ws.Range("Q2").Value = maxinc ' set the value in Q2 to the max value


     
    
    
    'Find Greatest % decrease
    Dim maxdec As Double
    Dim decTicker As String
    
    maxdec = WorksheetFunction.Min(maxRange)
    decTicker = tickerRange.Cells(Application.WorksheetFunction.Match(maxdec, maxRange, 0)).Value
    
    ws.Range("P3").Value = decTicker
    ws.Range("Q3").Value = maxdec
    
    
    
    
    'Find greatest total volume
    Dim LastRowTot As Long
    Dim Total As Double
    Dim TotTicker As String
    Dim TotRange As Range
    Dim MaxTot As Range
    
    LastRowTot = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    Set TotRange = ws.Range("I2:I" & LastRowTot)
    Set MaxTot = ws.Range("l2:L" & LastRowTot)
    
    Total = WorksheetFunction.Max(MaxTot)
    TotTicker = TotRange.Cells(Application.WorksheetFunction.Match(Total, MaxTot, 0)).Value
    
    ws.Range("P4").Value = TotTicker
    ws.Range("Q4").Value = Total
    
    
    
    
   
    
    
    
    
    
    
    'Add Conditionals
    
    ' conditionals Macro
'

'
    Dim LastRowCon As Long
    Dim Conditional As Range
    LastRowCon = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    Set Conditional = ws.Range("J2:J" & LastRowCon)
' add two conditions to the range
    With Conditional.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = vbRed ' set the formatting if the first condition is met
    End With
    
    With Conditional.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = vbGreen ' set the formatting if the second condition is met
    End With


   
    
    
'Conditionals for Percent Columns
    Dim LastRowPerCon As Long
    Dim PerCon As Range
    LastRowPerCon = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    Set PerCon = ws.Range("K2:K" & LastRowPerCon)
    
    With PerCon.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
        .Interior.Color = vbRed
    End With
    
    With PerCon.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="0")
        .Interior.Color = vbGreen
    End With
    

'Go to the next worksheet
    
    Next ws
    

End Sub
