VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Market()

Dim i As Long
Dim YearlyChange As Double
Dim totsum As Long

Dim Vol As Variant
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim PercentChange As Double
Dim LastRow As Long
Dim gi As Variant
Dim git As String
Dim gd As Variant
Dim gdt As String
Dim gtv As Variant
Dim gtvt As String




For Each ws In Worksheets
ws.Cells(1, 9) = "Ticker"
ws.Cells(1, 10) = "Yearly" + " " + "Change"
ws.Cells(1, 11) = "Percent" + " " + "Change"
ws.Cells(1, 12) = "Total" + " " + "Stock" + "Volume"
ws.Cells(2, 14) = "Greatest" + " " + "%" + "Increase"
ws.Cells(3, 14) = "Greatest" + " " + "%" + "decrease"
ws.Cells(4, 14) = "Greatest" + " " + "Total" + "Volume"
ws.Cells(1, 15) = "Ticker"
ws.Cells(1, 16) = "Value"

OpeningPrice = ws.Cells(2, 3).Value
PercentChange = 0
YearlyChange = 0
Volume = 0
ClosingPrice = 0
totsum = 2
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
gi = 0
gd = 999999999
gtv = 0

For i = 2 To LastRow


Vol = Vol + ws.Cells(i, 7).Value
    
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
    ws.Cells(totsum, 9).Value = ws.Cells(i, 1).Value
    ws.Cells(totsum, 12).Value = Vol
    ClosingPrice = ws.Cells(i, 6).Value
    
    YearlyChange = ClosingPrice - OpeningPrice

    
    

    ws.Cells(totsum, 10).Value = YearlyChange
    
    
        If YearlyChange <= 0 Then
        
            ws.Cells(totsum, 10).Interior.ColorIndex = 3
        Else: ws.Cells(totsum, 10).Interior.ColorIndex = 4
        
        End If
        
        
        
        If OpeningPrice > 0 Then
            
            PercentChange = YearlyChange / OpeningPrice
        
            
        Else: PercentChange = 0
        
        End If
        
    ws.Cells(totsum, 11).Value = PercentChange
            
    
    
    
    
    OpeningPrice = ws.Cells(i + 1, 3).Value
    totsum = totsum + 1
    
   Vol = 0
   
   End If
   
   
    If ws.Cells(i, 12).Value > gtv Then
    
    gtv = ws.Cells(i, 12).Value
    gtvt = ws.Cells(i, 9).Value
    
    End If
   
   If ws.Cells(i, 11).Value > gi Then
   gi = ws.Cells(i, 11).Value
   git = ws.Cells(i, 9).Value
   End If
   
   If ws.Cells(i, 10).Value < gd Then
   gd = ws.Cells(i, 10).Value
   gdt = ws.Cells(i, 9).Value
   End If


Next i

ws.Cells(4, 16).Value = gtv
ws.Cells(4, 15).Value = gtvt
ws.Cells(2, 16).Value = gi
ws.Cells(2, 15).Value = git
ws.Cells(3, 16).Value = gd
ws.Cells(3, 15).Value = gdt

 
Next ws

End Sub




