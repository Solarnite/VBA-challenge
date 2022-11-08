Sub StockMarket()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim percent As Double
Dim greatest_inc As Double
Dim greatest_dec As Double
Dim greatest_vol As Double
Dim start As Integer

For Each ws In Worksheets

    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"

    ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("P3").NumberFormat = "0.00%"

    start = 2
    prior = 1
    vol = 0
    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            ticker = ws.Cells(i, 1).Value

            prior = prior + 1

            year_open = ws.Cells(prior, 3).Value
            year_close = ws.Cells(i, 6).Value

            For j = prior To i

                vol = vol + ws.Cells(j, 7).Value

            Next j

            If year_open = 0 Then

                percent = year_close

            Else
                yearly_change = year_close - year_open

                percent = yearly_change / year_open

            End If
            
            ws.Cells(start, 9).Value = ticker
            ws.Cells(start, 10).Value = yearly_change
            ws.Cells(start, 11).Value = FormatPercent(percent)
            ws.Cells(start, 12).Value = vol
                    
            If ws.Cells(start, 10) > 0 Then
    
                ws.Cells(start, 10).Interior.ColorIndex = 4
    
            Else
    
                ws.Cells(start, 10).Interior.ColorIndex = 3
                    
            End If
            
            start = start + 1

            vol = 0
            yearly_change = 0
            percent = 0
            prior = i
            
        End If


    Next i
    
    greatest_inc = 0
    greatest_dec = 0
    greatest_vol = 0
    LastRow_K = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For k = 2 To LastRow_K
    
            current = ws.Cells(k, 11).Value
            volume = ws.Cells(k, 12).Value
            
            If current > greatest_inc Then
    
                greatest_inc = current
                ws.Range("P2").Value = greatest_inc
                ws.Range("P2").NumberFormat = "0.00%"
                ws.Range("O2").Value = ws.Cells(k, 9).Value
    
            End If
            
            If current < greatest_dec Then
    
                greatest_dec = current
                ws.Range("P3").Value = greatest_dec
                ws.Range("P3").NumberFormat = "0.00%"
                ws.Range("O3").Value = ws.Cells(k, 9).Value
    
            End If
            
            If volume > greatest_vol Then
    
                greatest_vol = volume
                ws.Range("P4").Value = greatest_vol
                ws.Range("O4").Value = ws.Cells(k, 9).Value
    
            End If
            
        Next k

Columns("I:Q").AutoFit
    
Next ws
End Sub
