sub Module2Challenge ()
    for each ws in Worksheets
        'Find Last Row
        LastRow=ws.cells(Rows.Count,1).End(xlUp).Row
        'Set up table
            Dim TickerCount as Integer
            TickerCount = 2
            Dim Ticker as String
            Dim TotalStock as double
            OpenValue = ws.range("C2").value
            'table headers Insert
            ws.Range("I1").value = "Ticker"
            ws.Range("J1").value = "Yearly Change"
            ws.Range("K1").value = "Percent Change"
            ws.Range("L1").value = "Total Stock Volume"
            'loop through all tickers
            for i = 2 to LastRow
                'identify ticker names
                if ws.cells(i,1).value <> ws.cells(i+1,1).value Then
                Ticker = ws.Cells(i, 1).Value
                'input ticker name in Table
                ws.cells(TickerCount,9).value = Ticker
                    'calculate yearly change (placeholder)
                    Dim YearlyChange as double
                    CloseValue=ws.cells(i,6).value
                    YearlyChange = CloseValue-OpenValue
                        'calculate Percent Change (placeholder)
                        Dim PercentChange as double
                        if OpenValue<>0 Then
                        PercentChange = YearlyChange / OpenValue
                        Else PercentChange = 0
                        End If
                        'ws.cells(i,11).NumberFormat = "0.00%"
                        'input percent change into table
                        ws.cells(TickerCount,11).value = PercentChange
                    OpenValue=ws.cells(i+1,3).value
                        if YearlyChange>=0 Then 
                        ws.Cells(TickerCount,10).Interior.ColorIndex=4
                        End If
                        if YearlyChange<0 Then
                        ws.Cells(TickerCount,10).Interior.ColorIndex=3
                        End If
                    '  input yearly change into table
                    ws.cells(TickerCount,10).value = YearlyChange
                        
                            'calculate Total Stock Volume
                            TotalStock = TotalStock + ws.Cells(i, 7).Value
                            'input total stock change change into table
                            ws.cells(TickerCount,12).value = TotalStock
                            TickerCount=TickerCount+1
                            TotalStock=0
                Else TotalStock = TotalStock + ws.Cells(i, 7).Value
                End if
            Next i

    'autofit colums for each ws because not being able o read the headers was driving me crazy
        ws.Range("I1:L1").EntireColumn.AutoFit

                        'headers for summary table
                            ws.Range("Q2").Value = "Greatest % increase"
                            ws.Range("Q3").Value = "Greatest % Decrease"
                            ws.Range("Q4").Value = "Greatest Total Volume"
                            ws.Range("R1").Value = "Ticker"
                            ws.Range("S1").Value = "Value"
                            GreatInc= ws.cells(i,11).value
                                for i = 2 to LastRow
                                    if ws.cells(i,11)>GreatInc Then
                                    GreatInc = ws.cells(i,11).value
                                    GreatIncTick = ws.cells(i,9).value
                                    End if
                                next i 
                            ws.Range("S2").value=GreatInc
                            ws.Range("R2").value = GreatIncTick

                                GreatDec= ws.cells(i,11).value
                                for i = 2 to LastRow
                                    if ws.cells(i,11)<GreatDec Then
                                    GreatDec = ws.cells(i,11).value
                                    GreatDecTick = ws.cells(i,9).value
                                    End if
                                next i 
                            ws.Range("S3").value=GreatDec
                            ws.Range("R3").value = GreatDecTick

                                GreatVol= ws.cells(i,12).value
                                for i = 2 to LastRow
                                    if ws.cells(i,12)>GreatVol Then
                                    GreatVol = ws.cells(i,12).value
                                    GreatVolTick = ws.cells(i,9).value
                                    End if
                                next i 
                            ws.Range("S4").value=GreatVol
                            ws.Range("R4").value = GreatVolTick

                            
        ws.Range("Q2:Q5").EntireColumn.AutoFit
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("S2:S3").NumberFormat = "0.00%"
    Next ws
    MsgBox("We did it, Joe!")
End Sub