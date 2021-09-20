Attribute VB_Name = "Module1"
Sub StockSummaryTables():

Dim sheetCount, rowcount, newIndex As Integer
Dim openVal, closeVal, annualGain, annualPercent, totVol, gGain, gLoss, gVol As Double
Dim symbol, symG, symL, symV As String

    sheetCount = Sheets.Count

    For i = 1 To sheetCount
  
        newIndex = 2
        totVol = 0
        gGain = 0
        gLoss = 0
        gVol = 0

        Worksheets(Sheets(i).Name).Activate
        rowcount = ActiveSheet.UsedRange.Rows.Count
        
        openVal = Cells(2, 3).Value
        Cells(1, 10).Value = "Ticker"
        Cells(1, 11).Value = "Yearly Change"
        Cells(1, 12).Value = "Percent Change"
        Cells(1, 13).Value = "Total Stock Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        For j = 2 To rowcount
    
            If Cells(j, 1).Value = Cells(j + 1, 1).Value Then
                totVol = totVol + Cells(j, 7).Value
                symbol = Cells(j, 1).Value
            Else
                totVol = totVol + Cells(j, 7).Value
                closeVal = Cells(j, 6).Value
                annualGain = closeVal - openVal
                If openVal <> 0 Then
                    annualPercent = annualGain / openVal
                Else
                    annualPercent = 0
                End If
                
                Cells(newIndex, 10).Value = symbol
                Cells(newIndex, 11).Value = annualGain
                Cells(newIndex, 12).Value = annualPercent
                Cells(newIndex, 12).NumberFormat = "0.00%"
                Cells(newIndex, 13).Value = totVol
                
                If annualGain >= 0 Then
                    Cells(newIndex, 11).Interior.ColorIndex = 4
                    Cells(newIndex, 12).Interior.ColorIndex = 4
                Else
                    Cells(newIndex, 11).Interior.ColorIndex = 3
                    Cells(newIndex, 12).Interior.ColorIndex = 3
                End If
                
                If annualPercent > gGain Then
                    gGain = annualPercent
                    symG = symbol
                ElseIf annualPercent < gLoss Then
                    gLoss = annualPercent
                    symL = symbol
                End If
                If totVol > gVol Then
                    gVol = totVol
                    symV = symbol
                End If
                
                totVol = 0
                newIndex = newIndex + 1
                openVal = Cells(j + 1, 3).Value
            End If
                
        
        Next j
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(2, 16).Value = symG
        Cells(3, 16).Value = symL
        Cells(4, 16).Value = symV
        Cells(2, 17).Value = gGain
        Cells(2, 17).NumberFormat = "0.00%"
        Cells(3, 17).Value = gLoss
        Cells(3, 17).NumberFormat = "0.00%"
        Cells(4, 17).Value = gVol
        
        Worksheets(Sheets(i).Name).Columns("J:M").EntireColumn.AutoFit
        Worksheets(Sheets(i).Name).Columns("O:Q").EntireColumn.AutoFit
    
    Next i

End Sub



