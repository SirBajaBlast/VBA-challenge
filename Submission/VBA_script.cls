VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stocks()

    Dim ws As Worksheet
    Dim i As Long ' row number
    Dim cell_vol As LongLong ' contents of column G (cell volume)
    Dim vol_total As LongLong ' what is going to go in column L
    Dim ticker As String ' what is going to go in column I
    Dim k As Long ' leaderboard row
    Dim ticker_close As Double
    Dim ticker_open As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim lastRow As Long
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As LongLong
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate ' Activate the worksheet

        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        vol_total = 0
        k = 2

        ' Write Leaderboard Columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Volume"

        ' assign open for first ticker
        ticker_open = ws.Cells(2, 3).Value

        For i = 2 To lastRow
            cell_vol = ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value

            If (ws.Cells(i + 1, 1).Value <> ticker) Then
                vol_total = vol_total + cell_vol

                ticker_close = ws.Cells(i, 6).Value
                price_change = ticker_close - ticker_open

                If (ticker_open > 0) Then
                    percent_change = price_change / ticker_open
                Else
                    percent_change = 0
                End If

                ws.Cells(k, 9).Value = ticker
                ws.Cells(k, 10).Value = price_change
                ws.Cells(k, 11).Value = percent_change
                ws.Cells(k, 12).Value = vol_total

                ' Formatting
                If (price_change > 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 4 ' Green
                ElseIf (price_change < 0) Then
                    ws.Cells(k, 10).Interior.ColorIndex = 3 ' Red
                Else
                    ws.Cells(k, 10).Interior.ColorIndex = 2 ' White
                End If

                vol_total = 0
                k = k + 1
                ticker_open = ws.Cells(i + 1, 3).Value
            Else
                vol_total = vol_total + cell_vol
            End If
        Next i

        ' Style the leaderboard
        ws.Columns("K:K").NumberFormat = "0.00%"
        ws.Columns("I:L").AutoFit

        ' Label the worksheet as Q1, Q2, Q3, or Q4
        Select Case ws.Index
            Case 1
                ws.Name = "Q1"
            Case 2
                ws.Name = "Q2"
            Case 3
                ws.Name = "Q3"
            Case 4
                ws.Name = "Q4"
        End Select
        
    ' Initialize variables
    maxIncrease = 0
    maxDecrease = 0
    maxVolume = 0
    maxIncreaseTicker = ""
    maxDecreaseTicker = ""
    maxVolumeTicker = ""
    
    ' Find the greatest percent change increase, decrease, and total volume
        For i = 2 To lastRow
            If ws.Cells(i, "L").Value > maxVolume Then
                maxVolume = ws.Cells(i, "L").Value
                maxVolumeTicker = ws.Cells(i, "I").Value
            End If

            If ws.Cells(i, "K").Value > maxIncrease Then
                maxIncrease = ws.Cells(i, "K").Value
                maxIncreaseTicker = ws.Cells(i, "I").Value
            End If

            If ws.Cells(i, "K").Value < maxDecrease Then
                maxDecrease = ws.Cells(i, "K").Value
                maxDecreaseTicker = ws.Cells(i, "I").Value
            End If
        Next i

        ' Output the results in rows O2, O3, and O4 with new column labels in P1 and Q1
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("Q2").Value = maxIncrease

        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("Q3").Value = maxDecrease

        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        
        ' Style the leaderboard
        Range("Q2").NumberFormat = "0.00%"
        Range("Q3").NumberFormat = "0.00%"
        ws.Columns("O:O").AutoFit
        
    Next ws

End Sub



