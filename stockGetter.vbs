Sub stockGetter():
    Dim solutionSheet As Worksheet
    Dim LastCol As Long
    Dim LastRow As Long
    Dim currentTicker As String
    Dim stockOpen As Double
    Dim stockClose As Double
    Dim i As Long
    Dim currentVolume As Double
    Dim Year As String
    Dim wb As Workbook
    Dim GreatestIncrease As Double
    Dim GreatestDecrease As Double
    Dim MaxVolume As Double
    Dim GreatestIncreaseTicker As String
    Dim GreatestDecreaseTicker As String
    Dim MaxVolumeTicker As String
    
    Set wb = ActiveWorkbook

    For Each ws In ActiveWorkbook.sheets
        Year = ws.Name
        
        ' Reset Solution Sheet
        DeleteSolutionSheet (Year + " - Solved")
        Set solutionSheet = wb.sheets.Add(Type:=xlWorksheet, After:=ws)
        solutionSheet.Name = Year + " - Solved"
        solutionSheet.Cells(1, 1).Value = "Ticker"
        solutionSheet.Cells(1, 2).Value = "Yearly Change"
        solutionSheet.Cells(1, 3).Value = "Pct. Yearly Change"
        solutionSheet.Cells(1, 4).Value = "Total Stock Volume"
        LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' set Variables
        currentTicker = ""
        GreatestIncreaseTicker = ""
        GreatestDecreaseTicker = ""
        MaxVolumeTicker = ""
        stockOpen = 0
        stockClose = 0
        currentVolume = 0
        GreatestIncrease = 0
        GreatestDecrease = 0
        MaxVolume = 0
        j = 2
        
        ' Loop to lastrow + 1 to get final stock
        For i = 2 To LastRow + 1

            If currentTicker <> ws.Cells(i, 1).Value Then ' We have a new stock!
                'skip calculations for the first one. these get run when we find the next one
                If i <> 2 Then
                    stockClose = ws.Cells(i - 1, 6).Value
                    solutionSheet.Cells(j, 1).Value = currentTicker
                    solutionSheet.Cells(j, 2).Value = stockClose - stockOpen
                    
                    ' Set color, leave it alone if there is no change
                    If stockClose > stockOpen Then
                        solutionSheet.Cells(j, 2).Interior.Color = RGB(0, 250, 0)
                    ElseIf stockClose < stockOpen Then
                        solutionSheet.Cells(j, 2).Interior.Color = RGB(250, 0, 0)
                    End If
                    
                    solutionSheet.Cells(j, 4).Value = currentVolume
                    
                    ' check for new max volume
                    If currentVolume > MaxVolume Then
                        MaxVolume = currentVolume
                        MaxVolumeTicker = currentTicker
                    End If
                    

                    ' catch division by zero 
                    If stockOpen > 0 Then
                        solutionSheet.Cells(j, 3).Value = Str(Round((stockClose - stockOpen) / stockOpen * 100, 2)) + "%"

                        ' check for new greatestest increase or decrease, using an elseif in the same conditional because it can't be both (unless there's only one stock)
                        If ((stockClose - stockOpen) / stockOpen * 100) > GreatestIncrease Then
                            GreatestIncrease = ((stockClose - stockOpen) / stockOpen * 100)
                            GreatestIncreaseTicker = currentTicker
                        ElseIf ((stockClose - stockOpen) / stockOpen * 100) < GreatestDecrease Then
                            GreatestDecrease = ((stockClose - stockOpen) / stockOpen * 100)
                            GreatestDecreaseTicker = currentTicker
                        End If
                    Else
                        solutionSheet.Cells(j, 3).Value = "Undefined"
                    End If
                    
                    j = j + 1
                End If
                currentTicker = ws.Cells(i, 1).Value
                stockOpen = ws.Cells(i, 3).Value
                currentVolume = 0
            End If
            
            currentVolume = currentVolume + ws.Cells(i, 7).Value
        
        Next i

        solutionSheet.Range("H2").Value = "Greatest % Increase"
        solutionSheet.Range("H3").Value = "Greatest % Decrease"
        solutionSheet.Range("H4").Value = "Highest Volume"
        
        solutionSheet.Range("I2").Value = GreatestIncreaseTicker
        solutionSheet.Range("I3").Value = GreatestDecreaseTicker
        solutionSheet.Range("I4").Value = MaxVolumeTicker
        
        solutionSheet.Range("J2").Value = Str(Round(GreatestIncrease, 2)) + "%"
        solutionSheet.Range("J3").Value = Str(Round(GreatestDecrease, 2)) + "%"
        solutionSheet.Range("J4").Value = MaxVolume
    Next
End Sub

Sub DeleteSolutionSheet(Name)
    Application.DisplayAlerts = False
    Worksheets(Name).Delete
    Application.DisplayAlerts = True
End Sub
