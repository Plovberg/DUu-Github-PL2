VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ticker_Summary()
    Dim Ticker As String
    Dim LastRow As Long
    Dim AnnStkRow As Long 'row where ticker is found
    
    For Each ws In ThisWorkbook.Worksheets
        
        ws.Range("I1").Value = "Ticker" 'name the header
        
        AnnStkRow = 2 'Initializing AnnStkRow to the starting row 2 for each ws
        'loop thru rows in column A
        For i = 2 To LastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                'store the ticker symbol in column I
                ws.Cells(AnnStkRow, 9).Value = Ticker
            'move to the next row in column I
                AnnStkRow = AnnStkRow + 1
            End If
        Next i
    Next ws
        
End Sub

Sub Annual_Price_Change()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim AnnStkRow As Long
    Dim i As Long ' Initialize i for the loop
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set the title row for the output
        ws.Range("J1").Value = "Annual Price Change"
        ' Initialize AnnStkRow
        AnnStkRow = 2
        ' Find the last row with data in column A
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each row of the worksheet
        For i = 2 To LastRow
            ' search date for "0102"
            If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                
                ' Find the closing price for the same ticker on YYYY1231
                Dim RowIndex As Long
                For RowIndex = i To LastRow
                    If Right(ws.Cells(RowIndex, 2).Value, 4) = "1231" Then
                        ClosingPrice = ws.Cells(RowIndex, 6).Value
                        Exit For
                    End If
                Next RowIndex
                
                ' Calculate the annual price change
                Dim AnnPrcChange As Double
                AnnPrcChange = ClosingPrice - OpeningPrice
                
                ' Store the annual price change in column J
                ws.Cells(AnnStkRow, 10).Value = AnnPrcChange
                If AnnPrcChange > 0 Then
                
                ws.Cells(AnnStkRow, 10).Interior.ColorIndex = 4 'green for positive returns
                Else: ws.Cells(AnnStkRow, 10).Interior.ColorIndex = 3 'red for negative retuns
                
                End If
                
                ' Store the ticker in column I
                ws.Cells(AnnStkRow, 9).Value = Ticker
                
                ' Move to the next row
                AnnStkRow = AnnStkRow + 1
            End If
        Next i
    Next ws
End Sub

Sub SubTotalVolumeByYear()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim Volume As Double
    Dim TotalVolume As Double
    Dim OutputRow As Long
    Dim AnnualVolRange As Range
    Dim i As Long ' Initialize i for the loop
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize total volume
        TotalVolume = 0
        
        ' Find the last row with data in column A
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set the range for yearly volume output (exclude the first row for the header)
        Set AnnualVolRange = ws.Range("L2:L" & LastRow)
        
        ' Set the column header for annual volume if it's not already set
        If ws.Cells(1, 12).Value <> "Annual Volume" Then
            ws.Cells(1, 12).Value = "Annual Volume"
        End If
        
        ' Reset output row for each worksheet
        OutputRow = 1
        
        ' Loop through each row of the worksheet
        For i = 2 To LastRow ' Start from row 2
            ' Extract the ticker and volume from the current row
            Ticker = ws.Cells(i, 1).Value
            Volume = ws.Cells(i, 7).Value
            
            ' If it's a new ticker, output total volume and reset for the next ticker
            If Ticker <> ws.Cells(i - 1, 1).Value Then
                ' Output total volume for the previous ticker if it's not the first row
                If i > 2 Then
                    AnnualVolRange.Cells(OutputRow).Value = TotalVolume
                    OutputRow = OutputRow + 1 ' Move to the next output row
                End If
                
                ' Reset total volume for the next ticker
                TotalVolume = Volume
            Else
                ' If it's not a new ticker, accumulate volume
                TotalVolume = TotalVolume + Volume
            End If
        Next i
        
        ' Output total volume for the last ticker
        AnnualVolRange.Cells(OutputRow).Value = TotalVolume
    Next ws
End Sub
 

Sub Annual_Percent_Change()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim AnnStkRow As Long
    Dim i As Long ' Initialize i for the loop
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Set the column header for annual percent change if it's not already set
        If ws.Cells(1, 11).Value <> "Annual Percent Change" Then
            ws.Cells(1, 11).Value = "Annual Percent Change"
        End If
        
        ' Initialize AnnStkRow
        AnnStkRow = 2
        
        ' Find the last row with data in column A
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Loop through each row of the worksheet
        For i = 2 To LastRow
            ' Check if the date is YYYY0102
            If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                
                ' Find the closing price for the same ticker on YYYY1231
                Dim RowIndex As Long
                For RowIndex = i To LastRow
                    If Right(ws.Cells(RowIndex, 2).Value, 4) = "1231" Then
                        ClosingPrice = ws.Cells(RowIndex, 6).Value
                        Exit For ' Exit the loop once closing price is found
                    End If
                Next RowIndex
                
                ' Calculate the annual percent change (percentage change)
                Dim AnnPctChange As Double
                AnnPctChange = ((ClosingPrice - OpeningPrice) / OpeningPrice) * 100
                ' Round off output to two significant digits
                AnnPctChange = Round(AnnPctChange, 2)
                
                ' Store the annual percent change in column K
                ws.Cells(AnnStkRow, 11).Value = AnnPctChange ' Changed output column to K
                
                ' Store the ticker in column I
                ws.Cells(AnnStkRow, 9).Value = Ticker
                
                ' Move to the next row
                AnnStkRow = AnnStkRow + 1
            End If
        Next i
    Next ws
End Sub

Sub FindGreatestPercentageIncrease()
    Dim ws As Worksheet
    Dim maxIncrease As Double
    Dim maxIncreaseTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the greatest percentage increase
        maxIncrease = WorksheetFunction.Max(ws.Range("K:K"))
        maxIncreaseTicker = ws.Cells(Application.Match(maxIncrease, ws.Range("K:K"), 0), 9).Value
        
        ' Output the greatest percentage increase and its corresponding ticker
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("Q2").Value = maxIncrease
        
        ' Set the headers in cells P1 and Q1
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Set the value for cell O2
        ws.Range("O2").Value = "Greatest % Increase"
    Next ws
End Sub

Sub FindGreatestPercentageDecrease()
    Dim ws As Worksheet
    Dim maxDecrease As Double
    Dim maxDecreaseTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the greatest percentage decrease
        maxDecrease = WorksheetFunction.Min(ws.Range("K:K"))
        maxDecreaseTicker = ws.Cells(Application.Match(maxDecrease, ws.Range("K:K"), 0), 9).Value
        
        ' Output the greatest percentage decrease and its corresponding ticker
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("Q3").Value = maxDecrease
        
        ' Set the headers in cells P1 and Q1
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Set the value for cell O3
        ws.Range("O3").Value = "Greatest % Decrease"
    Next ws
End Sub

Sub FindGreatestTotalVolume()
    Dim ws As Worksheet
    Dim maxVolume As Double
    Dim maxVolumeTicker As String
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Find the greatest total volume
        maxVolume = WorksheetFunction.Max(ws.Range("L:L"))
        maxVolumeTicker = ws.Cells(Application.Match(maxVolume, ws.Range("L:L"), 0), 9).Value
        
        ' Output the greatest total volume and its corresponding ticker
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q4").Value = maxVolume
        
        ' Set the headers in cells P1 and Q1
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        ' Set the value for cell O4
        ws.Range("O4").Value = "Greatest Total Volume"
    Next ws
End Sub
