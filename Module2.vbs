Attribute VB_Name = "Module2"
Sub StocksSummary()

    For Each ws In Worksheets
    
    ws.Activate
    
     Dim Worksheetname As String
     Dim Ticker As String
     Dim YearlyChange As Double
     Dim PercentChange As Double
     Dim Volume As Double
     Dim LastRow As Double
     
     Dim VolumeTotal As Double
         
     Dim tablerow As Integer
         
     
     LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
     tablerow = 2
         
    
     Cells(1, 9).Value = "Ticker"
     Cells(1, 10).Value = "Yearly Change"
     Cells(1, 11).Value = "Percentage Change"
     Cells(1, 12).Value = "Total Stock Volume"
     
         For i = 2 To LastRow
             Volume = Cells(i, 7).Value
             
             If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                 VolumeTotal = VolumeTotal + Volume
                 
             Else
                 Ticker = Cells(i, 1).Value
             
                 VolumeTotal = VolumeTotal + Volume
                 
                         
                 Range("L" & tablerow).Value = VolumeTotal
                 Range("I" & tablerow).Value = Ticker
                 
                 tablerow = tablerow + 1
                 
                 VolumeTotal = 0
                 
              End If
                        
              
         Next i
    
    
    Next ws


End Sub
