' Loop for all sheets to call the code for one sheet
Sub StockAllSheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Stock
    Next
    Application.ScreenUpdating = True
End Sub
' Loop for one sheet's calculations
Sub Stock()
    Dim ticker As String
    Dim change As Double
    change = 0
    Dim totalvolume As Double
    totalvolume = 0
    Dim tickerrow As Integer
    tickerrow = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    Cells(1, 9).Value = "ticker"
    Cells(1, 10).Value = "yearly change"
    Cells(1, 11).Value = "percent change"
    Cells(1, 12).Value = "total stock volume"
    
    
    For I = 2 To lastrow
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
            ticker = Cells(I, 1).Value
            change = change + (Cells(I, 6).Value - Cells(I, 3).Value)
            totalvolume = totalvolume + Cells(I, 7).Value
            Cells(tickerrow, 9).Value = ticker
            Cells(tickerrow, 10).Value = change
                If Cells(tickerrow, 10).Value > 0 Then
                    Cells(tickerrow, 10).Interior.ColorIndex = 4
                Else
                     Cells(tickerrow, 10).Interior.ColorIndex = 3
                End If
               
                If Cells(I, 3).Value = 0 Then
                    Cells(tickerrow, 11).Value = 0
                Else
                    Cells(tickerrow, 11).Value = FormatPercent(change / Cells(I, 3), 2)
                End If
                
            Cells(tickerrow, 12).Value = totalvolume
            tickerrow = tickerrow + 1
            change = 0
            totalvolume = 0
        Else
            change = change + (Cells(I, 6).Value - Cells(I, 3).Value)
            totalvolume = totalvolume + Cells(I, 7).Value
            If Cells(I, 3).Value = 0 Then
                Cells(tickerrow, 11).Value = 0
            Else
                Cells(tickerrow, 11).Value = FormatPercent(change / Cells(I, 3), 2)
            End If 
        End If
    Next I
End Sub
