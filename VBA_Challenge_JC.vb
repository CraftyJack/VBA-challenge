Sub Stocktastic()

' declare some variables
Dim current As Worksheet
Dim vTicker As String
Dim vDate As Double
Dim vOpen As Double
Dim vClose As Double
Dim vTotalVolume As Double
Dim vSummary_Table_Row As Integer
' declare some more variables
Dim vMaxPctIncrease As Double
Dim vMaxPctIncTicker As String
Dim vMaxPctDecrease As Double
Dim vMaxPctDecTicker As String
Dim vMaxVolume As Double
Dim vMaxVolTicker As String

' For each worksheet in the workbook
For Each current In Worksheets
    ' Activate the current Worksheet
    current.Activate
    
    ' Stake out an area for the summary table
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    ' Set the first row for summary table entries
    vSummary_Table_Row = 2
      
    ' Read in the opening price for the first symbol
    vOpen = Cells(2, 3).Value
    'MsgBox ("Opening price on sheet " & current.Name & " is " & vOpen)
    
    ' Loop through all of the rows except the first one
    For i = 2 To (current.Cells(Rows.Count, 2).End(xlUp).Row)
        ' Check to see if the ticker symbol has changed
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ' If it has...
            ' Set the ticker symbol and print it to the summary table
            vTicker = Cells(i, 1).Value
            Cells(vSummary_Table_Row, 9).Value = vTicker
            ' Add to the total volume
            vTotalVolume = vTotalVolume + Cells(i, 7).Value
            ' Print the total volume to the summary table
            Cells(vSummary_Table_Row, 12).Value = vTotalVolume
            ' Clear the total volume
            vTotalVolume = 0
            ' Get the year's closing price
            vClose = Cells(i, 6)
            ' Calculate the year's price change using (vClose - vOpen) and write it to the table
            Cells(vSummary_Table_Row, 10).Value = (vClose - vOpen)
            ' Check to see if the price change is positive
            If (vClose - vOpen) > 0 Then
                ' if so, color the cell green
                Cells(vSummary_Table_Row, 10).Interior.ColorIndex = 4
            ' Check to see if the price change is negative
            ElseIf (vClose - vOpen) < 0 Then
                ' if so, color the cell red
                Cells(vSummary_Table_Row, 10).Interior.ColorIndex = 3
            End If
            ' Calculate the year's percent price change using (vClose-vOpen) / vOpen and write it to the table
            ' ...but first, we need to protect ourselves against stocks that open at zero.
            If vOpen = 0 Then
                ' MsgBox (vTicker & " opened the year at 0. Percent change cannot be calculated!")
                Cells(vSummary_Table_Row, 11).Value = 0
            Else
                Cells(vSummary_Table_Row, 11).Value = FormatPercent(((vClose - vOpen) / vOpen), 2)
            End If
            ' Get the opening price for the next ticker symbol
            vOpen = Cells(i + 1, 3).Value
            ' Increment the Summary Row by 1
            vSummary_Table_Row = vSummary_Table_Row + 1
        ' Otherwise add to the total volume
        Else
           vTotalVolume = vTotalVolume + Cells(i, 7).Value
        End If
    Next i

    ' Stake out an area for the super summary table
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    ' Zero out the variables for the super summary
    vMaxPctIncrease = 0
    vMaxPctDecrease = 0
    vMaxVolume = 0

    ' Now, let's go through our summary table to find the max increase, max decrease, and max volume
    For j = 2 To (current.Cells(Rows.Count, 9).End(xlUp).Row)
        ' Read in the percent change and compare it to vMaxPctIncrease
        If Cells(j, 11).Value > vMaxPctIncrease Then
            ' if it's bigger, it's the new vMaxPctIncrease.
            vMaxPctIncrease = Cells(j, 11).Value
            vMaxPctIncTicker = Cells(j, 9).Value
        ' Now check it against the decrease
        ElseIf Cells(j, 11).Value < vMaxPctDecrease Then
            ' if it's smaller, it's the new vMaxPctDecrease
            vMaxPctDecrease = Cells(j, 11).Value
            vMaxPctDecTicker = Cells(j, 9).Value
        End If
        ' Now look for the maximum volume
        If Cells(j, 12).Value > vMaxVolume Then
            vMaxVolume = Cells(j, 12).Value
            vMaxVolTicker = Cells(j, 9).Value
        End If
    Next j
    
    ' Now fill in the super summary table
    Cells(2, 16).Value = vMaxPctIncTicker
    Cells(2, 17).Value = FormatPercent(vMaxPctIncrease, 2)
    Cells(3, 16).Value = vMaxPctDecTicker
    Cells(3, 17).Value = FormatPercent(vMaxPctDecrease, 2)
    Cells(4, 16).Value = vMaxVolTicker
    Cells(4, 17).Value = vMaxVolume
Next
End Sub