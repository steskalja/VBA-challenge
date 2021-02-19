Sub Stock_Data()
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Variant
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrese As Double
    Dim greatestVolume As Double
    Dim sRange As Range
    Dim fRange As Range
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    
    j = 1
    total = 0
    change = 0
    start = 2

    
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To rowCount
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            total = total + Cells(i, 7).Value
            
            
            
            change = (Cells(i, 6) - Cells(start, 3))
            If change <> 0 And Cells(start, 3) <> 0 Then
                percentChange = Round((change / Cells(start, 3) * 100), 2)
            End If
            
            
            Cells(j + 1, "I").Value = Cells(i, 1).Value
            Cells(j + 1, "L").Value = total
            Cells(j + 1, "J").Value = change
            If change < 0 Then
                Cells(j + 1, "J").Interior.ColorIndex = 3
            Else
               Cells(j + 1, "J").Interior.ColorIndex = 4
            End If
                
            Cells(j + 1, "K").Value = percentChange
            
            
            j = j + 1
            
            
            total = 0
            
            
            start = i + 1
            
        Else
            total = total + Cells(i, 7).Value
        End If
    Next i
    
    rowCount = Cells(Rows.Count, "K").End(xlUp).Row
    'Gets Greatest Increase
    Set sRange = Range("K2:K" + LTrim(Str(rowCount)))
    Range("Q2") = WorksheetFunction.Max(sRange)
    Set fRange = sRange.Find(What:=Range("Q2").Value)
    Range("P2") = Cells(fRange.Row, fRange.Column - 2)
    'Gets Greatest Decrease
    Range("Q3") = WorksheetFunction.Min(sRange)
    Set fRange = sRange.Find(What:=Range("Q3").Value)
    Range("P3") = Cells(fRange.Row, fRange.Column - 2)
    'Gets Greatest Volume
    Set sRange = Range("L2:L" + LTrim(Str(rowCount)))
    Range("Q4") = WorksheetFunction.Max(sRange)
    Set fRange = sRange.Find(What:=Range("Q4").Value)
    Range("P4") = Cells(fRange.Row, fRange.Column - 3)
    
    
    
    
End Sub

Sub Sheet_Info()
    'Cycles through Sheets and manipulates data
    For Each W In Worksheets
        W.Activate
        Stock_Data
        
    Next W
End Sub

