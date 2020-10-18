Sub NEW_stock()
'loop though the worksheets
Dim wks As Worksheet


For Each wks In worksheets

                wks.Cells(1, 9).Value = "Ticker"
                wks.Cells(1, 16).Value = "Ticker"
                wks.Cells(1, 10).Value = "Yearly Change"
                wks.Cells(1, 11).Value = "Percent Change"
                wks.Cells(1, 12).Value = "Total Stock Volume"
                wks.Cells(1, 17).Value = "value"
                wks.Cells(2, 15).Value = "Greatest % Increase"
                wks.Cells(3, 15).Value = "Greatest % Decrease"
                wks.Cells(4, 15).Value = "Greatest Total Volume"
    Next wks
    
   'declare var. calculations
   
Dim i As Long
Dim tickerName As String
Dim openYearly  As Double
Dim totalVolume As Double
totalVolume = 0
Dim totalYearly As Double
totalYearly = 0
Dim percentChange As Double
Dim tickerRow As Long
tickerRow = 2
Dim lastRow As Integer




'add loop

For i = 2 To lastRow
lastRow = wks.Cells(Rows.Count, "A").End(xlUp).Row

openYearly = ws.Cells(tickerRow, 3).Value

'add conditional
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    tickerName = ws.Cells(i, 1).Value
                    
                    ws.Range("I" & tickerRow).Value = tickerName
                    
                    totalYearly = totalYearly + (ws.Cells(i, 6).Value - openYearly)
                    ws.Range("J" & tickerRow).Value = totalYearly
                    
                    
                    percentChange = (totalYearly / openYearly)
                    ws.Range("K" & tickerRow).Value = percentChange
                    ws.Range("K" & tickerRow).Style = "Percent"
                    
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                    ws.Range("L" & tickerRow).Value = totalVolume
                    
                    'reset
                    
                    tickerRow = tickerRow + 1
                    totalYearly = 0
                    totalVolume = 0
                    openYearly = wks.Cells(tickerRow, 3).Value
                Else
                    totalVolume = totalVolume + ws.Cells(i, 7).Value
                End If
            Next i
            
            'declare var for formattiong
            
            Dim yearLastRow As Long
          
        
           
            
            'add loop for formatting
            
  For i = 2 To yearLastRow
  yearLastRow = wks.Cells(Rows.Count, "10").End(xlUp).Row
            
            'adding formatting conditionals
                    If wks.Cells(i, 10).Value >= 0 Then
                        wks.Cells(i, 10).Interior.ColorIndex = 4
                    Else
                        wks.Cells(i, 10).Interior.ColorIndex = 3
     End If
  Next i
            
            'declare Variables for finding max & min
            
            Dim percentLastRow As Long
            percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
            Dim percent_max As Double
            percent_max = 0
            Dim percent_min As Double
            percent_min = 0
            
            
    'add loop for max & min
        For i = 2 To percentLastRow
        
    'add conditionl for max and min
        If percent_max < wks.Cells(i, 11).Value Then
            percent_max = wks.Cells(i, 11).Value
            wks.Cells(2, 17).Value = percent_max
            wks.Cells(2, 17).Style = "percent"
            wks.Cells(2, 16).Value = wks.Cells(i, 9).Value
        ElseIf percent_min > wks.Cells(i, 11).Value Then
                percent_min = wks.Cells(i, 11).Value
                wks.Cells(3, 17).Value = percent_min
                wks.Cells(3, 17).Style = "Percent"
                wks.Cells(3, 16).Value = wls.Cells(i, 9).Value
        End If
    Next i
    
'variable for greatest vol

            Dim totalVolumeRow As Long
            totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
            Dim totalVolumeMax As Double
            totalVolumeMax = 0
            
    'loop for greatest total volume
        For i = 2 To totalVolumeRow
        
        'conditionals for greatwst total volume
            If totalVolumeMax < wks.Cells(i, 12).Value Then
                totalVolumeMax = wks.Cells(i, 12).Value
                wks.Cells(4, 17).Value = totalVolumeMax
                wks.Cells(4, 16).Value = wks.Cells(i, 9).Value
            End If
        Next i
       
        
  

End Sub