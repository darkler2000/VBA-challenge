Attribute VB_Name = "Module1"
Sub Module2Challenge()

' Define variables
Dim ws As Worksheet
Dim lastRow As Long
Dim ticker As String
Dim oPrice As Double
Dim cPrice As Double
Dim yearChange As Double
Dim perChange As Double
Dim totalVol As Double
Dim aggregate As Integer
Dim greatPerInc As Double
Dim greatPerDec As Double
Dim greatTotVol As Double
Dim greatPerIncTicker As String
Dim greatPerDecTicker As String
Dim greatTotVolTicker As String


For Each ws In Worksheets

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

ws.Range("K:K").NumberFormat = "0.00%"


' Define lastRow
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

' Initalize variable
aggregate = 2

For i = 2 To lastRow

    
    ' First entry business
    If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
        
        ' Set ticker
        ticker = ws.Cells(i, 1).Value
        
        'Set Opening Price
        oPrice = ws.Cells(i, 3).Value
        
        ' Initialize variable
        totalVol = 0
        
    End If
    
    ' Add to total volume
    totalVol = totalVol + ws.Cells(i, 7).Value
    
    'Last entry business
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        'Set Closing Price
        cPrice = ws.Cells(i, 6).Value
    
   
    
        ' Set Yearly Change
        yearChange = cPrice - oPrice
    
        ' Set Percent Change
        If oPrice = 0 Then
            perChange = 0
        Else
            perChange = yearChange / oPrice
        End If
    
    
    
        ' Populate column headers
        ws.Cells(aggregate, 9).Value = ticker
        ws.Cells(aggregate, 10).Value = yearChange
        ws.Cells(aggregate, 11).Value = perChange
        ws.Cells(aggregate, 12).Value = totalVol
        

        ' Color Cells
        If yearChange >= 0 Then
            ws.Cells(aggregate, 10).Interior.Color = RGB(0, 255, 0)
        Else
            ws.Cells(aggregate, 10).Interior.Color = RGB(255, 0, 0)
        End If
        
        aggregate = aggregate + 1
    
    End If
    
    

        


Next i

' Reinitalize variable
aggregate = 2
greatPerInc = 0
greatPerDec = 0
greatTotVol = 0

' Populate labels
ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"
ws.Cells(2, 14).Value = "Greatest % Increase"
ws.Cells(3, 14).Value = "Greatest % Decrease"
ws.Cells(4, 14).Value = "Greatest Total Volume"


For j = 2 To lastRow
    
    'Greatest % Increase
    If ws.Cells(j, 11).Value > greatPerInc Then
        greatPerInc = ws.Cells(j, 11).Value
        greatPerIncTicker = ws.Cells(j, 9).Value
    End If
    
    'Greatest % Decrease
    If ws.Cells(j, 11).Value < greatPerDec Then
        greatPerDec = ws.Cells(j, 11).Value
        greatPerDecTicker = ws.Cells(j, 9).Value
    End If
    
    'Greatest Total Volume
    If ws.Cells(j, 12).Value > greatTotVol Then
        greatTotVol = ws.Cells(j, 12).Value
        greatTotVolTicker = ws.Cells(j, 9).Value
    End If
    
    'Populate cells
    ws.Cells(2, 16).Value = greatPerInc
    ws.Cells(3, 16).Value = greatPerDec
    ws.Cells(4, 16).Value = greatTotVol
    ws.Cells(2, 15).Value = greatPerIncTicker
    ws.Cells(3, 15).Value = greatPerDecTicker
    ws.Cells(4, 15).Value = greatTotVolTicker
    ws.Range("p2:p3").NumberFormat = "0.00%"
    
Next j
        


Next ws



End Sub


