Attribute VB_Name = "Module1"
Sub StockData()

'Define worksheet as a worksheet object variable
Dim Ws As Worksheet

'Loop through each worksheet while encapsulating entire block of code for each worksheet loop
For Each Ws In Worksheets

'Define Lastrow to calculate the last row in each worksheet
Dim Lastrow As Double

'Define summary table row counter
Dim RowCounter As Integer

'Define row counter for OpenPrice
Dim OpenPriceRow As Double

'Define Opening and Closing Price variables
Dim OpenPrice As Double
Dim ClosePrice As Double

'Define summary output value variables
Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStockVolume As Double








'Initialize counters to begin at row 2
OpenPriceRow = 2
RowCounter = 2

'Ws.Activate borrowed from TA Farshad to fix error in Lastrow calculation
Ws.Activate

'Find the last row of each worksheet
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Checked to see if each worksheet is looped through.
'MsgBox (Ws.Name)

'Checked to see if last row of each worksheet is found.
'MsgBox (Lastrow)

'Reset the total stock volume for each worksheet
TotalStockVolume = 0

'Loop through each worksheet to output Ticker, Yearly Change, Percent Change, and Total Stock Volume
For i = 2 To Lastrow
    
    'Sum the total stock volume for each ticker
    TotalStockVolume = TotalStockVolume + Ws.Cells(i, 7)
    
    'If ticker value changes, then perform the code beneath to determine summary values for each ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        'When ticker value changes, assign it to Ticker
        Ticker = Ws.Cells(i, 1).Value
        
        'Store Ticker value in column I based on RowCounter increment
        Ws.Range("I" & RowCounter).Value = Ticker
    
        'Store open and close price; use those values to calculate yearly change, then store yearly change in column J
        
        OpenPrice = Ws.Range("C" & OpenPriceRow).Value
        ClosePrice = Ws.Range("F" & i).Value
        YearlyChange = ClosePrice - OpenPrice
        Ws.Range("J" & RowCounter).Value = YearlyChange
        
        'Assign total stock volume to column L
        Ws.Range("L" & RowCounter).Value = TotalStockVolume
        
            'Discussed with Tutor. Check for division by 0, then calculate percent change
            If OpenPrice <> 0 Then
                PercentChange = (YearlyChange / OpenPrice) * 100
        
            End If
        
        'Assign percent change to column K as percentage and round to two decimal places
        Ws.Range("K" & RowCounter).Value = "%" & Round(PercentChange, 2)
        
                'Highly positive yearly change in green and negative yearly change in red
                If YearlyChange > 0 Then
            
                    Ws.Range("J" & RowCounter).Interior.ColorIndex = 4
            
                ElseIf YearlyChange < 0 Then
            
                    Ws.Range("J" & RowCounter).Interior.ColorIndex = 3
            
                End If
        
        
        'increment counters by 1
        RowCounter = RowCounter + 1
        OpenPriceRow = i + 1
        
        'reset variables before next loop iteration
        PercentChange = 0
        ClosePrice = 0
        YearlyChange = 0
        TotalStockVolume = 0
    
      End If
        
Next i

'reset open price to zero before moving to next worksheet
OpenPrice = 0


Next Ws

End Sub






