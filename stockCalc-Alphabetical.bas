Attribute VB_Name = "Module1"
Sub stockCalc()


'Set variables to be used
'string to save ticker
Dim ticker As String
'double to save volume total
Dim volTotal As Double
'double to save opening price
Dim openP As Double
'double to save percent change
Dim perCh As Double
'double to save difference of close price and open price
Dim diffP As Double
'int to save row number for results
Dim stRow As Integer
'ticker for greatest percent increase
Dim gtPIT As String
'value of greatest percent increase
Dim gtPIV As Double
'ticker for greatest percent decrease
Dim gtPDT As String
'value for greatest percent decrease
Dim gtPDV As Double
'ticker for greatest total volume
Dim gTVT As String
'value of gratest total volume
Dim gTVV As Double



'loop through each worksheet
For Each ws In Worksheets
    
    'get total rows of the current worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set the first ticker
    ticker = ws.Cells(2, 1).Value
    '0 out total
    volTotal = 0
    'set the first open price
    openP = ws.Cells(2, 3).Value
    'set the first close price
    closeP = ws.Cells(2, 6).Value
    
    'setup headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'setup labels
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'reset stRow at the beginning of each ws
    stRow = 2
    
    'reset all values at the beginning of each ws
    gtPIV = 0
    gtPDV = 0
    gTVV = 0
    
    
    'loop from 2nd row to lastrow
    For i = 2 To lastrow
        
        
        'Check if value in cell (i,1) is equal to ticker
        If ws.Cells(i, 1).Value = ticker Then
            'increase volTotal with value (i,1)
            volTotal = volTotal + ws.Cells(i, 7).Value
            
        Else
            'End of last ticker, start of new ticker. Output last ticker's results
            'set closeP with the last value for close price in (i-1,6) = points to last closeP of previous ticker
            closeP = ws.Cells(i - 1, 6).Value
            'calculate change from open price to close price
            diffP = closeP - openP
            'calculate percent change from open price to close price
            perCh = diffP / openP
            
            'output results of Ticker and Yearly Changein current results row - determined by stRow
            ws.Cells(stRow, 9).Value = ticker
            ws.Cells(stRow, 10).Value = diffP
            'format Yearly Change to currency(dollars)
            ws.Cells(stRow, 10).NumberFormat = "$#,##0.00"
    
            'format Yearly Change respectively
            If diffP >= 0 Then
                'make green if positive
                ws.Cells(stRow, 10).Interior.ColorIndex = 4
            Else
                'make red otherwise
                ws.Cells(stRow, 10).Interior.ColorIndex = 3
            End If
            
            
            'output results for Percent Change and Total Stock Volume
            ws.Cells(stRow, 11).Value = perCh
            'format Percent Change to percentage
            ws.Cells(stRow, 11).NumberFormat = "0.00%"
            ws.Cells(stRow, 12).Value = volTotal
            
            'check if current Percent Change is the greatest or the lowest overall
            If perCh > gtPIV Then
                'if greater, store value of Pecent Change in gtPIV
                gtPIV = perCh
                'store value of ticker to gtPIT
                gtPIT = ticker
                'ws.Cells(2, 15) = gtPIT
            ElseIf perCh < gtPDV Then
                'if lesser, store value of Percent Change in gtPDV
                gtPDV = perCh
                'store value of ticker to gtPDT
                gtPDT = ticker
                'ws.Cells(3, 15) = gtPDT
            End If
            
            'check if current volTotal is the greatest total overall
            If volTotal > gTVV Then
                'store ticker to gTVT
                gTVT = ticker
                'store total to gTVV
                gTVV = volTotal
                'ws.Cells(4, 15) = gTVT
            End If
            
            'increase stRow so next entry will be on the next row
            stRow = stRow + 1
            
            'change variables to store new entries for Ticker, Total Stock Volume, and Open Price
            ticker = ws.Cells(i, 1).Value
            volTotal = ws.Cells(i, 7).Value
            openP = ws.Cells(i, 3).Value
            
        End If
    
    Next i
    
            
            'Calculate Yearly Change and Percent Change of the last row of results
            'set closeP with the last value for close price in (i-1,6) = points to last closeP of previous ticker
            closeP = ws.Cells(i - 1, 6).Value
            'calculate change from open price to close price
            diffP = closeP - openP
            'calculate percent change from open price to close price
            perCh = diffP / openP
            
            
            'output last row values since the last row will not output in the loop
            ws.Cells(stRow, 9).Value = ticker
            ws.Cells(stRow, 10).Value = diffP
            ws.Cells(stRow, 10).NumberFormat = "$#,##0.00"
            ws.Cells(stRow, 11).Value = perCh
            ws.Cells(stRow, 11).NumberFormat = "0.00%"
            ws.Cells(stRow, 12).Value = volTotal
            
            'Check Percent Change if positive or negative and format accordingly
            If perCh >= 0 Then
                ws.Cells(stRow, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(stRow, 10).Interior.ColorIndex = 3
            End If
            
            
            'Check if final row results will change the stored gtPIV, gtPIT, gtPDV, gtPDT
            'output final results
            If perCh > gtPIV Then
                gtPIV = perCh
                gtPIT = ticker
                'output if last row has Greatest Percent Increase
                ws.Range("P2").Value = gtPIV
                'format to percentage
                ws.Cells(2, 16).NumberFormat = "0.00%"
                'output ticker
                ws.Range("O2").Value = gtPIT
            ElseIf perCh < gtPDV Then
                gtPDV = perCh
                gtPDT = ticker
                MsgBox "Changed " & gtPDT
                'output if last row has greatest Percent Decrease
                ws.Range("P3").Value = gtPDV
                'format to percentage
                ws.Cells(3, 16).NumberFormat = "0.00%"
                'output ticker
                ws.Range("O3").Value = gtPDT
            Else
                'output stored Greatest Percent Decresase and format as percentages
                ws.Cells(3, 16).Value = gtPDV
                ws.Cells(3, 16).NumberFormat = "0.00%"
                ws.Cells(3, 15).Value = gtPDT
                
                'output stored Percent Increase and format as percentages
                ws.Cells(2, 16).Value = gtPIV
                ws.Cells(2, 16).NumberFormat = "0.00%"
                ws.Cells(2, 15).Value = gtPIT
            End If
            
            'Check if final row volume is the Greatest Total Volume
            If volTotal > gTVV Then
                'output ticker
                ws.Range("O4").Value = gTVT
                gTVV = volTotal
                'output volume if Greatest Total Volume
                ws.Range("P4").Value = gTVV
            Else
                'output stored volume as Greatest Total Volume
                ws.Range("P4").Value = gTVV
                'output srored ticker of Greatest Total Volume
                ws.Range("O4").Value = gTVT
            End If
            
            'Autofit columns
            Worksheets(ws.Name).Columns("A:P").AutoFit
            
            
            
   
Next ws

'output that macro has finished
MsgBox "Complete!"



End Sub


