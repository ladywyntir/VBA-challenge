Attribute VB_Name = "Module11"
' Module 2 VBA Homework
'Created by Marjorie Muñoz
'
' Loop through all stocks for 1 year and output the following
'   - Ticker symbol
'   - Yearly change from opening price at year start to closing price at year end
'   - % of change from opening price to closing price
'   - Total stock volume
'
'BONUS
'Also calculate the following and show the Ticker Symbol and value for each
'   - Greatest % increase
'   - Greatest $ decrease
'   - Greatest total volume
    
'My default clear button command to be able to re-run my script
Sub Clear()

For Each ws In Worksheets
    ws.Range("I1:z40000").Value = ""            'removes content
    ws.Range("I1:z40000").ClearFormats          'removes formating
    ws.Range("I1:z40000").ColumnWidth = 8.43    'sets columns back to their default width
Next ws

End Sub

'my Stock Ticker subscript
Sub tickerLoop():
    
    'looping through all worksheets
    Dim currentWs As Worksheet      'variable for the current worksheet
    
    For Each currentWs In Worksheets
    
    'variable declarations
    Dim tickerSymbol As String      'variable for ticker symbol
    Dim lastRow As Long             'variable for last row
    Dim totalStockVol As LongLong   'variable for stock volume running total
    Dim i As Long                   'variable for the row used in the for loop
    Dim soy_openPrice As Double     'variable for start of year open price
    Dim eoy_closedPrice As Double   'variable for end of year closed price
    Dim yearlyChange As Double      'variable for the yearly change (delta)
    Dim percentChange As Double     'variable for the percent of change
    
    'Bonus calculation variable declarations
    Dim greatPerInc As Double       'variable for the greatest % increase
    Dim greatPerDec As Double       'variable for the greatest % decrease
    Dim greatTtlVol As LongLong     'variable for the greatest total volume
    Dim greatTkrInc As String       'variable for the ticker with greatest increase
    Dim greatTkrDec As String       'variable for the ticker with greatest decrease
    Dim greatTkrTtlVol As String    'variable for the ticker name with the greatest volume
    
    'variable initial values
    tickerSymbol = ""       'ticker symbol starts empty
    tableRow = 2            'table starter row starts at 2 since row 1 is the header
    totalStockVol = 0       'stock volume running total starts at 0
    soy_openPrice = 0       'start of year open price starts at 0
    eoy_closedPrice = 0     'end of year closed price starts at 0
    yearlyChange = 0        'yearly change starts at 0
    percentChange = 0       'percent change starts at 0
        
        
    'Bonus variable initial values
    greatPerInc = 0         'greatest percent increase starts at 0
    greatPerDec = 0         'greatest percent decrease starts at 0
    greatTtlVol = 0         'greatest ticket volume starts at 0
    greatTkrInc = ""        'greatest increase ticker symbol starts empty
    greatTkrDec = ""        'greatest decrease ticker symbol starts empty
    greatTkrTltvol = ""     'greatest volume ticker symbol starts empty
           
        'Calculating the last row
        lastRow = currentWs.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Setting the headers for the spreadsheet, columns I-L and then O-P
        currentWs.Range("I1:L1").Font.Bold = True             'bold font
        currentWs.Range("I1").Value = "Ticker Symbol"
        currentWs.Range("J1").Value = "Yearly Change"
        currentWs.Range("K1").Value = "Percent Change"
        currentWs.Range("L1").Value = "Total Stock Volume"
        currentWs.Range("O2").Value = "Greatest % Increase"
        currentWs.Range("O3").Value = "Greatest % Decrease"
        currentWs.Range("O4").Value = "Greatest Total Volume"
        currentWs.Range("P1").Value = "Ticker"
        currentWs.Range("Q1").Value = "Value"
        
        'Setting the initial start of year open price for the first ticker symbol
        soy_openPrice = currentWs.Cells(2, 3).Value
        
        'Looping through all rows
        For i = 2 To lastRow
            
            'Check for ticker symbol change:
            If currentWs.Cells(i + 1, 1).Value <> currentWs.Cells(i, 1).Value Then
            
                'If there is a change:
            
                'Set the values to the variables
                tickerSymbol = currentWs.Cells(i, 1).Value
                eoy_closedPrice = currentWs.Cells(i, 6).Value
                totalStockVol = totalStockVol + currentWs.Cells(i, 7).Value
                
                
                'MATH--------------------------------------------------
                yearlyChange = eoy_closedPrice - soy_openPrice          'Calculate the yearly change (delta)
                percentChange = (yearlyChange / soy_openPrice) * 100    'Calculate the percent of yearly change
                '------------------------------------------------------
                
                'Display Summaries-------------------------------------
                currentWs.Cells(tableRow, 9).Value = tickerSymbol                   'tickerSymbol in column I
                currentWs.Cells(tableRow, 10).Value = yearlyChange                  'yearlyChange in column J
                currentWs.Cells(tableRow, 11).Value = (CStr(percentChange) & "%")   'percentChange in column K
                currentWs.Cells(tableRow, 12).Value = totalStockVol                 'total stock volume in column L
                '------------------------------------------------------
                
                
                'Format yearly change cells based on position from 0
                If (yearlyChange > 0) Then
                    currentWs.Range("J" & tableRow).Interior.ColorIndex = 4   'green for positive change
                ElseIf (yearlyChange <= 0) Then
                    currentWs.Range("J" & tableRow).Interior.ColorIndex = 3   'red for a negative change
                
                End If
                            
                'Go to the next row by adding 1 to the current row
                tableRow = tableRow + 1
                
                
                'Set the next ticker symbol's start of year open price
                soy_openPrice = currentWs.Cells(i + 1, 3).Value
                
                
                                                
                'Bonus calculations
                If (percentChange > greatPerInc) Then       'if the current % of change is greater than the greatest % increase
                    greatPerInc = percentChange                 'set the % in the increase variable
                    greatTkrInc = tickerSymbol                  'set the ticker symbol in the variable
                ElseIf (percentChange < greatPerDec) Then   'if the current % of change is less than the greatest % decrease
                    greatPerDec = percentChange                 'set the % in the decrease variable
                    greatTkrDec = tickerSymbol                  'set the ticker symbol in the variable
                End If
       
                If (totalStockVol > greatTltVol) Then       'if the current total stock volume is greater than the greatest total volume
                    greatTltVol = totalStockVol                 'set the total volume in the greatest total variable
                    greatTkrTtlVol = tickerSymbol               'set the ticker symbol in the variable
                End If
                
                'Reset the variable values so the next ticker symbols can be summarized
                yearlyChange = 0
                percentChange = 0
                totalStockVol = 0
                                
            Else
             'If there is no change to the ticker symbol, add total Stock Volume
            totalStockVol = totalStockVol + currentWs.Cells(i, 7).Value
            
            End If
            
                        
            
        Next i
               
            'Used this as a test to run through the Worksheets
            'MsgBox ("This is sheet " + currentWs.Name)
                
            'Print the Greatest totals to the right of our other summary columns
            currentWs.Range("P2").Value = greatTkrInc
            currentWs.Range("Q2").Value = (CStr(greatPerInc) & "%")
            currentWs.Range("P3").Value = greatTkrDec
            currentWs.Range("Q3").Value = (CStr(greatPerDec) & "%")
            currentWs.Range("P4").Value = greatTkrTtlVol
            currentWs.Range("Q4").Value = greatTltVol
            ' Autofit to display data
            currentWs.Columns("O:P").AutoFit
            currentWs.Columns("I:L").AutoFit
    
    Next currentWs

End Sub
