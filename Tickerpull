'This first part of the sub is meant to pull all the unique values (ticker names)
'This needs the data organaized both by ticker name and then by date, with values in the same place. Easy to do in excel
'but this can't be stuck on any sheet and expected to work. Fairly rudamentary.

Sub tickerpull():
    'This first part pulls all the unique tickers in the data
    'We have to subtract 1 for headers in numrow

    numrows = Application.CountA(Range("a:a"))
    Dim Unique As Integer

    Unique = 2
    For i = 2 To numrows:
       If Cells(i, 1).Value <> Cells((i - 1), 1).Value Then
            Cells(Unique, 10).Value = Cells(i, 1).Value
            Unique = Unique + 1
        End If
    Next i

    'Now that i think about it, there's probably a function that just pulls unique values. Oh well
    
    'This second part pulls all the necessary data and places it next to the ticker
    
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim PercentCng As Double
    Dim volume As Double
    
    EndDate = WorksheetFunction.Max(Range("b:b"))
    volume = 0
    Unique = 2
    For i = 2 To numrows:
        If Cells(i, 2) < Cells((i - 1), 2) Or Cells((i - 1), 2) = vbString Then
            OpenPrice = Cells(i, 3).Value
            volume = volume + Cells(i, 7).Value
            opendate = Cells(i, 2)
        ElseIf Cells(i, 2).Value = EndDate Then
        'I realize far too late, this is the simplest way to pull the beginning/end date
        
            ClosePrice = Cells(i, 6).Value
            
            'Finding and printing data
            Cells(Unique, 11).Value = (ClosePrice - OpenPrice)
            PercentCng = (ClosePrice - OpenPrice) / OpenPrice
            Cells(Unique, 12).Value = PercentCng
            volume = volume + Cells(i, 7).Value
            
            Cells(Unique, 13).Value = volume
                'adding color change
                If PercentCng >= 0 Then
                    Cells(Unique, 11).Interior.Color = vbGreen
                Else
                    Cells(Unique, 11).Interior.Color = vbRed
                End If
                
            Unique = Unique + 1
            volume = 0
        Else
            volume = volume + Cells(i, 7).Value

        End If
    Next i
    
    MaxPerc = WorksheetFunction.Max(Range("L:L"))
    totalcount = Application.CountA(Range("L:L"))
    MinPerc = WorksheetFunction.Min(Range("L:L"))
    Maxvol = WorksheetFunction.Max(Range("M:M"))
        
    For i = 2 To totalcount:
        If Cells(i, 12).Value = MaxPerc Then
            MaxRow = Cells(i, 12).Row
            MaxTicker = Cells(MaxRow, 10)
            Range("p2").Value = MaxTicker
            Range("q2").Value = MaxPerc
            'Test script("Max is in row " & MaxRow)
        ElseIf Cells(i, 12).Value = MinPerc Then
            MinRow = Cells(i, 12).Row
            MinTicker = Cells(MinRow, 10)
            Range("p3").Value = MinTicker
            Range("q3").Value = MinPerc
        End If
        If Cells(i, 13).Value = Maxvol Then
            MaxVolRow = Cells(i, 12).Row
            MaxVolTic = Cells(MaxVolRow, 10)
            Range("p4").Value = MaxVolTic
            Range("q4").Value = Maxvol
        End If
    Next i
    'Now, to add formats and titles!
    Range("L:L").NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    Range("J1").Value = "Ticker"
    Range("k1").Value = "Qtr. Change"
    Range("L1").Value = "% Change"
    Range("M1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("o2").Value = "Greatest % Increase"
    Range("o3").Value = "Greatest % Decrease"
    Range("o4").Value = "Greatest Total Volume"
End Sub
