Attribute VB_Name = "Module1"
Sub init():

Range("I1") = "Ticker"
Range("J1") = "Yearly Change"
Range("K1") = "Percent Change"
Range("L1") = "Total Stock Volume"
ActiveSheet.Columns("I:L").AutoFit

End Sub

Sub formatting():

    Dim redGreen As Range

    Set redGreen = Range(Range("J2"), Range("J2").End(xlDown))

    redGreen.FormatConditions.Delete
    
'Set conditional format

    Set Green = redGreen.FormatConditions.Add(xlCellValue, xlGreater, "=0")
    Set Red = redGreen.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
    With Green
    .Interior.ColorIndex = 4
   End With
   
    With Red
    .Interior.ColorIndex = 3
   End With

End Sub

Sub sheetsLoop():

Dim i, k As Integer
Dim total As LongLong
Dim figureOpen, figureClose, yearChange As Double
'Set initial values
i = 2
k = 2
total = 0
figureOpen = Cells(2, 3).Value

Do While Not IsEmpty(Cells(i, 1).Value)
    'Sum up all volumes
    total = total + Cells(i, 7).Value
    
    If Cells(i + 1, 1) <> Cells(i, 1) Then
    'Add new Ticker
        Cells(k, 9) = Cells(i, 1)
    'Calculate yearly change
        figureClose = Cells(i, 6)
        yearChange = figureClose - figureOpen
        Cells(k, 10) = yearChange
    'Calculate percent change and display
        Cells(k, 11) = FormatPercent(yearChange / figureOpen, 2)
    'Display total stock volume
        Cells(k, 12) = total
    'Reset figures
        figureOpen = Cells(i + 1, 3).Value
        total = 0
    
        k = k + 1
        
    End If

    i = i + 1
Loop

End Sub

Sub bonus():

Dim maxPer, minPer As Double
Dim maxTot As LongLong

Range("O2") = "Greatest % Increase"
Range("O3") = "Greatest % Decrease"
Range("O4") = "Greatest Total Volume"
Range("P1") = "Ticker"
Range("Q1") = "Value"

'Display Greatest % Increase
Call columeSort("K1", False)
Range("P2") = Range("I2")
Range("Q2").Value = FormatPercent(Range("K2").Value, 2)
'Display Greatest % Decrease
Call columeSort("K1", True)
Range("P3") = Range("I2")
Range("Q3").Value = FormatPercent(Range("K2").Value, 2)
'Display Greatest Total Volume
Call columeSort("L1", False)
Range("p4") = Range("I2")
Range("Q4") = Range("L2")

Call columeSort("I1", True)

ActiveSheet.Columns("O:Q").AutoFit

End Sub

Sub columeSort(title As String, seq As Boolean):

If seq Then
    Range("I:L").Sort Key1:=Range(title), _
                     Order1:=xlAscending, _
                     Header:=xlYes
Else
    Range("I:L").Sort Key1:=Range(title), _
                     Order1:=xlDescending, _
                     Header:=xlYes
End If
End Sub



Sub main():

Dim ws As Worksheet

    ' Loop through all sheets
    For Each ws In Worksheets
        ws.Activate
        init
        sheetsLoop
        formatting
        bonus

    Next ws

End Sub

