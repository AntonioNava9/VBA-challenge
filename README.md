# VBA-challenge
VBA Analyssis
test
 the code

Sub Challenge2_worksheets()

For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yaerly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Range("I1:M1").HorizontalAlignment = xlCenter
    ws.Range("I1:M1").Font.FontStyle = "Bold"
    ws.Range("I1:M1").Columns.AutoFit
    ws.Cells(2, 15).Value = "Greates % Change"
    ws.Cells(3, 15).Value = "Greates % Decrease"
    ws.Cells(4, 15).Value = "Greates Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Range("O2:O4").Columns.AutoFit
    ws.Range("O2:O4").Font.FontStyle = "Bold"
    ws.Range("P1:Q1").Columns.AutoFit
    ws.Range("P1:Q1").Font.FontStyle = "Bold"


    
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row
          
    Dim StockName As String
    
    Dim OpenPrice As Double
    OpenPrice = 0
    
    Dim ClosingPrice As Double
    ClosingPrice = 0
    
    Dim Price_variation As Double
    
    Dim Pchange  As Double
    
    Dim TotalVolume As Double
    TotalVolume = 0
    
    Dim capture_price As Boolean

    Dim DataRow As Integer
    DataRow = 2
    
    For i = 2 To FinalRow
    
        
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
    
    StockName = Cells(i, 1).Value
    TotalVolume = TotalVolume + Cells(i, 7).Value
    ClosingPrice = Cells(i, 6).Value
    Price_variation = ClosingPrice - OpenPrice
    Pchange = ClosingPrice / OpenPrice - 1
    
    
    ws.Cells(DataRow, 9).Value = StockName
    ws.Cells(DataRow, 10).Value = Price_variation
    If ws.Cells(DataRow, 10).Value < 0 Then
    ws.Cells(DataRow, 10).Interior.ColorIndex = 3
    Else
    ws.Cells(DataRow, 10).Interior.ColorIndex = 4
    End If
    ws.Cells(DataRow, 11).Value = Pchange
    ws.Cells(DataRow, 11).NumberFormat = "0.00%"

    ws.Cells(DataRow, 12).Value = TotalVolume
    
    TotalVolume = 0
    DataRow = DataRow + 1
    capture_price = False
    Price_variation = 0
    Pchange = 0
   
    
    Else
    
    TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    If capture_price = False Then
    OpenPrice = ws.Cells(i, 3).Value
    capture_price = True
    
    End If
   
    
    End If
    
    Next i
           
    Dim maxchange As Double
    maxchange = 0
    
    Dim minchange As Double
    minchange = 0
    
    Dim maxvolume As Double
    maxvolume = 0
    
    Dim maxchanget As String
    Dim minchanget As String
    Dim maxvolumet As String

For j = 2 To FinalRow
    
    If maxchange < ws.Cells(j, 11).Value Then
        maxchange = ws.Cells(j, 11).Value
        maxchanget = ws.Cells(j, 9).Value
        
    End If
    
    If minchange > ws.Cells(j, 11).Value Then
        minchange = ws.Cells(j, 11).Value
        minchanget = ws.Cells(j, 9).Value
        
    End If
    
    If maxvolume < ws.Cells(j, 12).Value Then
        maxvolume = ws.Cells(j, 12).Value
        maxvolumet = ws.Cells(j, 9).Value
        
    End If
        
    Next j
    
    
    ws.Cells(2, 16).Value = maxchanget
    ws.Cells(2, 17).Value = maxchange
    ws.Cells(3, 16).Value = minchanget
    ws.Cells(3, 17).Value = minchange
    ws.Cells(4, 16).Value = maxvolumet
    ws.Cells(4, 17).Value = maxvolume
    
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("I1:Q1").Columns.AutoFit
    ws.Range("Q2:Q3").Columns.AutoFit
    
    Next ws

End Sub
