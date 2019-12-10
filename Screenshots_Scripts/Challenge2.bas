Attribute VB_Name = "Module2"
Sub AllWorksheets_Stock_Data_Loop()

   Dim WS As Worksheet
   
   For Each WS In Worksheets
   WS.Activate
   
    
    Dim Lastrow1 As LongLong


Lastrow1 = Range("A1048576").End(xlUp).Row

For I = Lastrow1 To 2 Step -1
If Cells(I, 3) = "0" Then
Rows(I).EntireRow.Delete
End If
Next I


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

For I = 2 To Lastrow
    If Cells(I - 1, 1).Value <> Cells(I, 1).Value Then
    Volume = Volume + Cells(I, 7).Value
    YearOpen = Cells(I, 3).Value
    
    ElseIf Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        Ticker = Cells(I, 1).Value
        YearEnd = Cells(I, 6).Value
        Volume = Volume + Cells(I, 7)
        Range("I" & Summary_Table_Row).Value = Ticker
        Range("L" & Summary_Table_Row).Value = Volume
        Range("J" & Summary_Table_Row).Value = YearEnd - YearOpen
        Range("K" & Summary_Table_Row).Value = (YearEnd - YearOpen) / (YearOpen)
        Summary_Table_Row = Summary_Table_Row + 1
        Volume = 0
        
        Else
        Volume = Volume + Cells(I, 7).Value
   
    End If
    
Next I

Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Annual Percent Change"
Range("L1").Value = "Total Volume"

Dim newlastrow As Long
newlastrow = Cells(Rows.Count, "I").End(xlUp).Row


For I = 2 To newlastrow
    Cells(I, 11).NumberFormat = "0.00%"
    Cells(I, 10).NumberFormat = "0.00"
    If Cells(I, 10).Value < 0 Then
    Cells(I, 10).Interior.ColorIndex = 3
    ElseIf Cells(I, 10).Value > 0 Then
    Cells(I, 10).Interior.ColorIndex = 4
    Else: Cells(I, 10).Interior.ColorIndex = 0
    End If
    Next I
    
Dim Greatest_Increase As Double
Dim Greatest_Decrease As Double
Dim Most_Volume As LongLong
Greatest_Increase = 0
Greatest_Decrease = 0
Most_Volume = 0


For I = 2 To newlastrow
    If Cells(I, 11).Value >= Greatest_Increase Then
    Greatest_Increase = Cells(I, 11).Value
    Range("O2") = Cells(I, 9).Value
    End If
    
    If Cells(I, 11).Value <= Greatest_Decrease Then
    Greatest_Decrease = Cells(I, 11).Value
    Range("O3") = Cells(I, 9).Value
    End If
    
    If Cells(I, 12).Value > Most_Volume Then
    Most_Volume = Cells(I, 12).Value
    Range("O4") = Cells(I, 9).Value
    End If
    
Next I

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

  Next
    

End Sub
