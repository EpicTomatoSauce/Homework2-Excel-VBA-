Attribute VB_Name = "SubmissionFinal"
Sub stocks_analysis()

Dim ticker As String
Dim stockdate As Date
Dim stockopen As Double
Dim stockhigh As Double
Dim stocklow As Double
Dim stockclose As Double
Dim stockvol As Double
Dim summary_ticker As Double
Dim yearly_change As Double
Dim stockopenrow As Long
Dim pctchange As Double


'Define Worksheet
Dim ws As Worksheet

'Begin Loop
For Each ws In ThisWorkbook.Worksheets

ws.Activate

    'Initial Values
    summary_ticker = 2
    stockvol = 0
    yearly_change = 0
    stockopen = 0
    stockclose = 0
    stockopenrow = 2
    
    'Form Structure
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "% Change"
    Range("L1") = "Total Stock Volume"
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("O2") = "Greatest % Increase"
    Range("O3") = "Greatest % Decrease"
    Range("O4") = "Greatest Total Volume"
        
    'Find Ticker Symbols
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    lastcolumn = Cells(1, Columns.Count).End(xlToLeft).column
            
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Find Ticker
            ticker = Cells(i, 1).Value
                
            'Stock Open
            stockopen = Cells(stockopenrow, 3).Value
            stockopenrow = i + 1
            
            'Find closing value of stock
            stockclose = Cells(i, 6).Value
            
            'Print Ticker Value
            Cells(summary_ticker, 9).Value = ticker
                
            'Add to stockvol
            stockvol = stockvol + Cells(i, 7).Value
                
            'Print stockvol
            Cells(summary_ticker, 12).Value = stockvol
                
            'Calculate and print yearly change in stock
            yearly_change = stockclose - stockopen
            Cells(summary_ticker, 10).Value = yearly_change
            If yearly_change < 0 Then
                Cells(summary_ticker, 10).Interior.ColorIndex = 3
            ElseIf yearly_change > 0 Then
                Cells(summary_ticker, 10).Interior.ColorIndex = 4
            End If
                    
            'Calculate % change
            If stockopen = 0 Then
                pctchange = 0
                Else
                pctchange = (stockclose / stockopen) - 1
                Cells(summary_ticker, 11).Value = pctchange
                Cells(summary_ticker, 11).NumberFormat = "0.00%"
            End If
                                                                                    
            'Next summary line
            summary_ticker = summary_ticker + 1
                
            'Reset values
            stockvol = 0
                          
            'If ticker+1 = ticker, then do
        Else
            stockvol = stockvol + Cells(i, 7).Value
        
        End If
    Next i
    
    yearly_change = 0
    stockopen = 0
    stockclose = 0
    
    'Challenges
    'Look for greatest % increase and decrease and total volume
    
    Dim max As Double
    Dim maxvalue As Double
    Dim TestValueMax As Double
    Dim min As Double
    Dim TestValueMin As Double
    Dim greatesttotal As Double
    Dim TestValueTotal As Double
    
    lastrowchallenge = Cells(Rows.Count, 11).End(xlUp).Row
    
    max = Cells(2, 11).Value
    min = Cells(2, 11).Value
    greatesttotal = Cells(2, 12).Value
    
    For a = 2 To lastrowchallenge
        
    TestValueMax = Cells(a, 11).Value
        
        If TestValueMax > max Then
        
        max = TestValueMax
        
        Range("Q2").Value = TestValueMax
        Range("P2").Value = Cells(a, 9)
        Range("Q2").NumberFormat = "0.00%"
        
    End If
            
    Next a
    
    For b = 2 To lastrowchallenge
    
    TestValueMin = Cells(b, 11).Value
        
        If TestValueMin < min Then
        
        min = TestValueMin
        
        Range("Q3").Value = TestValueMin
        Range("P3").Value = Cells(b, 9)
        Range("Q3").NumberFormat = "0.00%"
        
    End If
            
    Next b
    
    For c = 2 To lastrowchallenge
    
    TestValueTotal = Cells(c, 12).Value
        
    If TestValueTotal > greatesttotal Then
        
        greatesttotal = TestValueTotal
        
        Range("Q4").Value = TestValueTotal
        Range("P4").Value = Cells(c, 9)
        Range("Q4").NumberFormat = "0"
            
    End If
            
    Next c
                
    'Autofit all columns in worksheets of workbooks
    Dim wsautofit As Worksheet
    For Each wsautofit In ThisWorkbook.Worksheets
    wsautofit.Cells.EntireColumn.AutoFit
    
    Next wsautofit

Next ws

End Sub



