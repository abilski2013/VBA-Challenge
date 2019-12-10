Attribute VB_Name = "Module1"
Sub Stock_Data()

Dim Lastrow1 As LongLong


Lastrow1 = Range("A1048576").End(xlUp).Row

For i = Lastrow1 To 2 Step -1
If Cells(i, 3) = "0" Then
Rows(i).EntireRow.Delete
End If
Next i


Dim Ticker As String
Dim YearlyChange As Double
Dim YearOpen As Double
Dim YearEnd As Double

Dim Percent_Change As Double

Dim Volume As LongLong
Volume = 0
Dim Lastrow As Long
Lastrow = Cells(Rows.Count, "A").End(xlUp).Row
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

For i = 2 To Lastrow
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    Volume = Volume + Cells(i, 7).Value
    YearOpen = Cells(i, 3).Value
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        YearEnd = Cells(i, 6).Value
        Volume = Volume + Cells(i, 7)
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = Volume
        Range("J" & Summary_Table_Row).Value = YearEnd - YearOpen
        Range("K" & Summary_Table_Row).Value = (YearEnd - YearOpen) / (YearOpen)
        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0
        
        Else
        Volume = Volume + Cells(i, 7).Value
   
    End If
    
Next i

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Annual Percent Change"
Range("L1").Value = "Total Volume"

Dim newlastrow As Long
newlastrow = Cells(Rows.Count, "I").End(xlUp).Row


For i = 2 To newlastrow
    Cells(i, 11).NumberFormat = "0.00%"
    Cells(i, 10).NumberFormat = "0.00"
    If Cells(i, 10).Value < 0 Then
    Cells(i, 10).Interior.ColorIndex = 3
    ElseIf Cells(i, 10).Value > 0 Then
    Cells(i, 10).Interior.ColorIndex = 4
    Else: Cells(i, 10).Interior.ColorIndex = 0
    End If
    Next i
    
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Most_Volume As LongLong



For i = 2 To newlastrow
    If Cells(i, 11).Value > Greatest_Increase Then
    Greatest_Increase = Cells(i, 11).Value
    Range("O2") = Cells(i, 9).Value
    End If
    
    If Cells(i, 11).Value < Greatest_Decrease Then
    Greatest_Decrease = Cells(i, 11).Value
    Range("O3") = Cells(i, 9).Value
    End If
    
    If Cells(i, 12).Value > Most_Volume Then
    Most_Volume = Cells(i, 12).Value
    Range("O4") = Cells(i, 9).Value
    End If
    
Next i

    Range("N2") = "Greatest Increase"
    Range("N3") = "Greatest Decrease"
    Range("N4") = "Most Volume"
    
    Range("P2") = Greatest_Increase
    Range("P2").NumberFormat = "0.00%"
    
    Range("P3") = Greatest_Decrease
    Range("P3").NumberFormat = "0.00%"
    
    Range("P4") = Most_Volume
    
    
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
    
    
    
End Sub

