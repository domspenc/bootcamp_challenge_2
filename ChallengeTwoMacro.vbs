Attribute VB_Name = "Module1"
Sub challengeTwo():

Dim ws As Worksheet

For Each ws In Worksheets

ws.Activate

Dim datarowcount As LongLong
datarowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row

Dim tablerowcount As LongLong
tablerowcount = ws.Cells(Rows.Count, "K").End(xlUp).Row

Dim ticker As String

Dim stockvolume As LongLong
stockvolume = 0

Dim counter As Integer
counter = 2

Dim openprice As Double
Dim closingprice As Double
Dim yearlychange As Double
Dim percentchange As Double

'---------------------------------------------------------------------
'create table headers

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

'---------------------------------------------------------------------
'populate new table within first loop
'includes grabbing tickers, calculating yearly change, change as a %, and total volume

    For i = 2 To datarowcount
    ticker = ws.Cells(i, 1).Value
    
        If stockvolume = 0 Then
        openprice = ws.Cells(i, 3).Value
        
        End If
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(counter, 9).Value = ticker
            
            closingprice = ws.Cells(i, 6).Value
            yearlychange = closingprice - openprice
            ws.Cells(counter, 10).Value = yearlychange
            
            percentchange = (yearlychange / openprice)
            ws.Cells(counter, 11).Value = percentchange
            
            stockvolume = stockvolume + ws.Cells(i, 7).Value
            ws.Cells(counter, 12).Value = stockvolume
        
            counter = counter + 1
            stockvolume = 0
        
        Else
        stockvolume = stockvolume + ws.Cells(i, 7).Value
                
    
    End If
     
     
Next i

'---------------------------------------------------------------------
'conditional formatting entries within second loop, to stop the fill command
'at end of table cell selection

For i = 2 To tablerowcount


' Yearly Change column...

    If ws.Cells(i, 10).Value < 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 10).Value >= 0 Then
    ws.Cells(i, 10).Interior.ColorIndex = 4
  
  End If

' Percent Change column...

    If ws.Cells(i, 11).Value < 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 3
    
    ElseIf ws.Cells(i, 11).Value >= 0 Then
    ws.Cells(i, 11).Interior.ColorIndex = 4
  
  End If
  

Next i

'---------------------------------------------------------------------
' convert percent change column to percent formatting

Dim p As LongLong

    For p = 2 To tablerowcount
    ws.Range("K" & p).NumberFormat = "0.00%"
    
    Next p
    
    

'---------------------------------------------------------------------
' BONUS SECTION
'---------------------------------------------------------------------


'create headers

ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'---------------------------------------------------------------------
'greatest increase and decrease percentages and greatest volume loop

Dim bonusticker As String
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim bonusvolume As LongLong

    For i = 2 To tablerowcount
    bonusticker = ws.Cells(i, 9).Value
    greatestIncrease = ws.Cells(i, 11).Value
    greatestDecrease = ws.Cells(i, 11).Value
    bonusvolume = ws.Cells(i, 12).Value

        If (ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K:K"))) Then
        ws.Range("P2") = bonusticker
        ws.Range("Q2") = greatestIncrease
  
        ElseIf (ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K:K"))) Then
        ws.Range("P3") = bonusticker
        ws.Range("Q3") = greatestDecrease
        
        ElseIf (ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L:L"))) Then
        ws.Range("P4") = bonusticker
        ws.Range("Q4") = bonusvolume
        
        End If
    
    
    Next i


'---------------------------------------------------------------------
' convert percent change column to percent formatting

Dim per As Integer

    For per = 2 To 3
    ws.Range("Q" & per).NumberFormat = "0.00%"
    
    Next per
 
        
'---------------------------------------------------------------------
' extra aesthetic formatting commands for worksheet readability

ws.Range("A1:Q1").Font.Bold = True
ws.Range("I1:L91").BorderAround (xlContinuous)
ws.Range("O1:Q4").BorderAround (xlContinuous)
ws.Columns("A:Q").AutoFit



Next ws

End Sub
