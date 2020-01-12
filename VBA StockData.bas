Attribute VB_Name = "Module1"
Sub StockData():
    
Dim Ws_Count As Worksheet
    For Each Ws_Count In ActiveWorkbook.Worksheets
    Ws_Count.Activate
    

        ' Add Heading for summary
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        
        'Create Variable to hold Value
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
        Dim Volume As Double
        Volume = 0
        Dim row As Double
        row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        'Set Initial Open Price
        Open_Price = Cells(2, Column + 2).Value
         
         ' Loop through rows
        For i = 2 To Ws_Count.Range("A1").End(xlUp).row
        
         'if we are still within the same ticker symbol, if it is not...
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
            
                ' Set Ticker name
                Ticker_Name = Cells(i, Column).Value
                Cells(row, Column + 8).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(i, Column + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(row, Column + 10).Value = Percent_Change
                    Cells(row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(i, Column + 6).Value
                Cells(row, Column + 11).Value = Volume
                ' Add one to the summary table row
                row = row + 1
                ' reset the Open Price
                Open_Price = Cells(i + 1, Column + 2)
                ' reset the Volume Total
                Volume = 0
            'if cells have the same ticker
            Else
                Volume = Volume + Cells(i, Column + 6).Value
            End If
        Next i
        
        ' Set the Cell Colors
        For J = 2 To Ws_Count.Cells(Rows.Count, Column + 8).End(xlUp).row
            If (Cells(J, Column + 9).Value > 0 Or Cells(J, Column + 9).Value = 0) Then
                Cells(J, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(J, Column + 9).Value < 0 Then
                Cells(J, Column + 9).Interior.ColorIndex = 3
            End If
        Next J

        
    Next Ws_Count
        
End Sub

