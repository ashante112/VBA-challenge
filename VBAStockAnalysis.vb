Sub stock_analysis():

'STEP 1: LOOP THROUGH ALL WORKSHEETS
    
    For Each ws In Worksheets
    
'STEP 2: SET COLUMN HEADERS FOR REPORT

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
'STEP 3: SET VARIABLES
        
        Dim ticker As String
        Dim tickercounter As Double
        tickercounter = 0
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
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
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through each Row
        For i = 2 To LastRow

                'Add the Ticker Symbol to Rows
                 tickercounter = tickercounter + ws.Cells(i, 7).Value

                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    ws.Range("I" & ReportRow).Value = ticker
                    ws.Range("L" & ReportRow).Value = tickercounter
                    tickercounter = 0
                    
                 'Add yearly change to report
                    OpenPrice = ws.Range("C" & PriorAmount)
                    ClosePrice = ws.Range("F" & i)
                    YearlyChange = ClosePrice - OpenPrice
                    ws.Range("J" & ReportRow).Value = YearlyChange
                    
                    'Add percent change to report
                    If OpenPrice = 0 Then
                        PercentChange = 0
                    Else
                        OpenPrice = ws.Range("C" & PriorAmount)
                        PercentChange = YearlyChange / OpenPrice
                    End If
                    
                    'Add One to Rowcounter
                    ReportRow = ReportRow + 1
                    PriorAmount = i + 1
                End If
    
        Next i

        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
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
        
        ' Format Double
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            
        ' Format Table
        ws.Columns("I:Q").AutoFit

    Next ws
    
End Sub
