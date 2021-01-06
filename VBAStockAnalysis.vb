Sub stock_analysis():

'LOOP THROUGH ALL WORKSHEETS
    
    For Each ws In Worksheets
    
'SET COLUMN HEADERS FOR REPORT

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'SET VARIABLES
        
        Dim ticker As String
        Dim tickercounter As Double
        tickercounter = 0
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim PriorAmount As Long
        PriorAmount = 2
        Dim LastRow As Long
        Dim LastColumn As Long
        Dim ReportRow As Long
        ReportRow = 2
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0
        
        'DETERMINE LAST ROW
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'LOOP THROUGH EACH ROW
        For i = 2 To LastRow

                'ADD TICKER TO ROWS
                 tickercounter = tickercounter + ws.Cells(i, 7).Value

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & ReportRow).Value = ticker
                    ws.Range("L" & ReportRow).Value = tickercounter
                    tickercounter = 0
                    
                    'ADD YEARLY CHANGE
                    
                    OpenPrice = ws.Range("C" & PriorAmount)
                    ClosePrice = ws.Range("F" & i)
                    YearlyChange = ClosePrice - OpenPrice
                    ws.Range("J" & ReportRow).Value = YearlyChange
                    
                    'ADD CONDITIONAL FORMATTING
                    
                    If YearlyChange > 0 Then
                
                    ws.Range("J" & ReportRow).Interior.Color = vbGreen
                    
                    ElseIf YearlyChange < 0 Then
                    
                    ws.Range("J" & ReportRow).Interior.Color = vbRed
                    
                    End If
                    
                    'ADD PERCENT CHANGE
                    
                    If OpenPrice = 0 Then
                        PercentChange = 0
                    
                    Else
                        OpenPrice = ws.Range("C" & PriorAmount)
                        PercentChange = YearlyChange / OpenPrice
                    
                    End If
                    
                ' Format Double To Include % Symbol And Two Decimal Places
                    ws.Range("K" & ReportRow).NumberFormat = "0.00%"
                    ws.Range("K" & ReportRow).Value = PercentChange

                    'ADD TO ROWCOUNTER
                    
                    ReportRow = ReportRow + 1
                    PriorAmount = i + 1
                
                End If
    
        Next i

        'SET CONDITIONS FOR REPORT
        
        For i = 2 To LastRow
        
            If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                ws.Range("Q2").Value = ws.Range("K" & i).Value
                ws.Range("P2").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                ws.Range("Q3").Value = ws.Range("K" & i).Value
                ws.Range("P3").Value = ws.Range("I" & i).Value
            End If

            If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                ws.Range("Q4").Value = ws.Range("L" & i).Value
                ws.Range("P4").Value = ws.Range("I" & i).Value
            End If

        Next i
        
        ' FORMAT DOUBLE
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
            
        ' FORMAT TABLE
        ws.Columns("I:Q").AutoFit

    Next ws
    
End Sub
