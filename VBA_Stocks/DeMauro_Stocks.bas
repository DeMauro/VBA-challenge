Attribute VB_Name = "DeMauro_Stocks"
Sub StocksHW()
'turn off screen updates to speed up macro
Application.ScreenUpdating = False

    'set variables
    Dim XTicker, NTicker, Ticker As String
    Dim lRow, lcol, row As Long
    Dim dteS, dteE, Max, Min, tv, tvs As Double
    Dim cell As Range
    Dim ws As Worksheet
    
'might as well start at the begining
Sheets(1).Select
'code to perform actions on all worksheets in workbook, with "Next ws." below
For Each ws In Worksheets
'selecting each sheet means I don't have to type ws. in front of everything (coverting from 1 sheet to all sheets)
ws.Select

    'LRow and row for looping, dteS is start date for stock, lcol isn't necessary but I used it for sorting
    lRow = Cells(Rows.Count, 1).End(xlUp).row
    lcol = Cells(1, Columns.Count).End(xlToLeft).Column
    dteS = Range("c2")
    row = 2
             
    'sort data to make sure stock symbols and dates are in order
    Range("A2", Cells(lRow, lcol)).Sort Key1:=Range("a2"), Order1:=xlAscending, _
    Key2:=Range("b1"), Order2:=xlAscending
    
    'add headers for columns and set number formats
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Chng"
    Range("K1") = "Percent Chng"
    Range("L1") = "Total Stock Volume"
    Range("J:J").NumberFormat = "0.00"
    Range("K:K").NumberFormat = "0.00%"
    Range("o2") = "Greatest % Increase"
    Range("o3") = "Greatest % Decrease"
    Range("o4") = "Greatest Total Volume"
    Range("p1") = "Ticker"
    Range("q1") = "Value"
    Range("q2:q3").NumberFormat = "0.00%"
    
    'loop through all stock symbols, perform calculations, and add data to columns
    For i = 2 To lRow
        If (Cells(i + 1, 1) <> Cells(i, 1)) Then
            'ticker symbol
            Range("I" & row) = Cells(i, 1)
            'Year Change and cell format '(red loss, green gain, no fill for no change)
            dteE = Cells(i, 6)
            Range("j" & row) = dteE - dteS
            If Range("j" & row) < 0 Then
                Range("j" & row).Interior.ColorIndex = 3 'red
            ElseIf Range("j" & row) = 0 Then
                Range("j" & row).Interior.ColorIndex = 0 'no fill
            Else: Range("j" & row).Interior.ColorIndex = 4 'green
            End If
            'percent change (with conditional for divide by zero). Might as well color it too
            If dteS = 0 Then
                Range("K" & row) = 0
                Else:  Range("k" & row) = (dteE / dteS) - 1
            End If
            If Range("k" & row) < 0 Then
                Range("k" & row).Interior.ColorIndex = 3 'red
            ElseIf Range("k" & row) = 0 Then
                Range("k" & row).Interior.ColorIndex = 0 'no fill
            Else: Range("k" & row).Interior.ColorIndex = 4 'green
            End If
            'total stock volume
            tvs = tvs + Cells(i, 7)
            Range("L" & row) = tvs
            'get starting stock price of next stock, reset total stock value, and move rows for data
            dteS = Cells(i + 1, 3)
            tvs = 0
            row = row + 1
        Else:
            'total stock volume
            tvs = tvs + Cells(i, 7)
        End If
    Next i
    
   'max and Min percent change (could have done "for i" but I've used this before)
    For Each cell In Range(Range("k2"), Range("k2").End(xlDown))
        If cell > Max Then
            Max = cell
            XTicker = cell.Offset(0, -2)
        End If
        If cell < Min Then
                Min = cell
                MTicker = cell.Offset(0, -2)
        End If
    Next cell
    Range("p2") = XTicker
    Range("q2") = Max
    Range("p3") = MTicker
    Range("q3") = Min
    Min = 0
    Max = 0
    'Greatest Total Volume
    For Each cell In Range(Range("L2"), Range("L2").End(xlDown))
        If cell > tv Then
            tv = cell
            Ticker = cell.Offset(0, -3)
        End If
    Next cell
    Range("p4") = Ticker
    Range("q4") = tv
    tv = 0
    Range("i:q").EntireColumn.AutoFit
Next ws

'return to the 1st worksheet and turn on screen updates
Sheets(1).Select
Application.ScreenUpdating = True

End Sub
