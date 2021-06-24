Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim i As Long
Dim YearlyChange As Double
Dim SummaryRow As Integer
Dim TotalVolume As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim PercentChange As Double

PercentChange = 0
YearlyChange = 0
TotalVolume = 0
OpeningPrice = Cells(2, 3).Value
ClosingPrice = 0
SummaryRow = 2
For i = 2 To 70926

    TotalVolume = TotalVolume + Cells(i, 7).Value
    
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    Cells(SummaryRow, 9).Value = Cells(i, 1).Value
    Cells(SummaryRow, 13).Value = TotalVolume
    
    ClosingPricing = Cells(i, 6).Value
    
    YearlyChange = ClosingPrice - OpeningPrice
    
    Cells(SummaryRow, 10).Value = YearlyChange
    
    
        If YearlyChange <= 0 Then
        
            Cells(SummaryRow, 10).Interior.ColorIndex = 3
        Else: Cells(SummaryRow, 10).Interior.ColorIndex = 4
        
        End If
        
        
        
        If OpeningPrice > 0 Then
            
            PercentChange = YearlyChange - OpeningPrice
            
        Else: PercentChange = 0
        
        End If
        
    Cells(SummaryRow, 12).Value = PercentChange
            
    OpeningPrice = Cells(i + 1, 3).Value
    
    
    SummaryRow = SummaryRow + 1
    
    TotalVolume = 0
    
    
  
    End If
    
    
    

Next i



End Sub

