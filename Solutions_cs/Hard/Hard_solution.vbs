Sub stock_analysis():

    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double

    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    ' Set initial values
    j = 0
    total = 0
    change = 0
    start = 2

    ' get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To rowCount

        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Stores results in variables
            total = total + Cells(i, 7).Value

            ' Handle zero total volume
            If total = 0 Then
                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0

            Else
                ' Find First non zero starting value
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                     Next find_value
                End If

                ' Calculate Change
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = Round((change / Cells(start, 3) * 100), 2)

                ' start of the next stock ticker
                start = i + 1

                ' print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(change, 2)
                Range("K" & 2 + j).Value = "%" & percentChange
                Range("L" & 2 + j).Value = total

                ' colors positives green and negatives red
                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 0
                End Select

            End If

            ' reset variables for new stock ticker
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
        Else
            total = total + Cells(i, 7).Value

        End If

    Next i

    ' take the max and min and place them in a separate part in the worksheet
    Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
    Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
    Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

    ' returns one less because header row not a factor
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

    ' final ticker symbol for  total, greatest % of increase and decrease, and average
    Range("P2") = Cells(increase_number + 1, 9)
    Range("P3") = Cells(decrease_number + 1, 9)
    Range("P4") = Cells(volume_number + 1, 9)

End Sub

Sub nextcell()
    Dim Stock_Symbol As String
    Dim Summary_Table_Row As Integer
    Dim Total_Volume As Double
    Dim Price_Change As Double
    Dim Current_Price As Double
    Dim Starting_Price As Double
    Dim Percent_Change As Double
    Dim ran As Range
    Dim d As Long
    Dim c As Double
    Set ran = Range("A2", Range("A2").End(xlDown))
    d = ran.Cells.Count + 1

    beginning = 2

    Cells(1, 9).Value = "Stock Symbol"
    Cells(1, 10).Value = "Year Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Summary_Table_Row = 2

    For i = 2 To d
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Stock_Symbol = Cells(i, 1).Value
            Total_Volume = Total_Volume + Cells(i, 7).Value
            Range("I" & Summary_Table_Row).Value = Stock_Symbol
            Range("L" & Summary_Table_Row).Value = Total_Volume
            Summary_Table_Row = Summary_Table_Row + 1
            Total_Volume = 0

            change = Cells(i, 6) - Cells(beginning, 3)
            'Here is where i'm setting the change into my new table - put code 
            'here to update the table
            'I'm using the current value of beginning, and then for the next ticker
            'I'll be using beginning again but it will be set to the beginning of that ticker, which 
            'is i + 1

            Else
             Percent_Change = "NA"
            beginning = i + 1
        Else
            Total_Volume = Total_Volume + Cells(i, 7).Value
        End If
    Next i

    Else
     Range("K" & Summary_Table_Row).Value = "NA"

    Dim rg As Range
    Dim r As Long
    Set rg = Range("J2", Range("J2").End(xlDown))
    r = rg.Cells.Count + 1
    For j = 2 To r
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        Else
            Cells(j, 10).Interior.ColorIndex = 2
        End If
    Next j

End Sub