Sub Ticker()

'Summary Table Creation
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

'Format PercentChange Column
Columns("K:K").Select
Selection.NumberFormat = "0.00%"
Columns("J:J").Select
Selection.NumberFormat = "0.00"

'Identifying Values
Dim Ticker As String
Dim i As Double
Dim OpenPrice As Double
Dim TableRow As Double
Dim TotalVolume As Double
Dim YearlyChange As Variant
Dim PercentChange As Variant
Dim j As Double

'Find Last Row
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
TableRow = 2
TotalVolume = 0
OpenPrice = 2
j = 2


    'For Loop
        For i = 2 To LastRow
    
        If Cells(i, 1) <> Cells(i + 1, 1) Then
        'EX: If Cells(7, 1) <> Cells(8, 1) Then

        'Table Generation
        Ticker = Cells(i, 1).Value
        'EX: Ticker = Cells(7, 1).Value
        OpenPrice = Cells(j, 3).Value
        'EX: OpenPrice = Cells (2, 3).Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        'EX: TotalVolume = 0 + Cells(7, 7).Value
        YearlyChange = Cells(i, 6).Value - OpenPrice
        'EX: YearlyChange = Cells(7, 6).Value - Cells(7, 3)
        
            'check for '0' value of open price
            If Cells(j, 3) = 0 Then
                For new_j = j To i
                    If Cells(new_j, 3).Value <> 0 Then
                        j = new_j
                        Exit For
                    End If
                Next new_j
            End If
        
        PercentChange = YearlyChange / OpenPrice
        'EX: PercentChange = Yearly Change / OpenPrice
        

          'Placing Values in Table
          Range("I" & TableRow) = Ticker
          Range("L" & TableRow) = TotalVolume
          Range("J" & TableRow) = YearlyChange
          Range("K" & TableRow) = PercentChange
    
          
          If YearlyChange > 0 Then
          Range("J" & TableRow).Interior.ColorIndex = 4
    
          Else
          Range("J" & TableRow).Interior.ColorIndex = 3
    
          End If
          
          'Reset Volume >> Next Ticker
          TableRow = TableRow + 1
          TotalVolume = 0
          j = i + 1
    
        Else
        TotalVolume = TotalVolume + Cells(i, 7).Value
           

        End If
    
        Next i

    End Sub
