Attribute VB_Name = "stockchange"
Sub stockyearlychange()
    ' Declaring variables
    Dim ticker As String
    Dim openstock As Double
    Dim closingstock As Double
    Dim lr As Long
    Dim ws As Worksheet
    Dim total As Double
    Dim output As Double

    ' Initializing loop for each worksheet using for and next ws
    For Each ws In Worksheets
        ' Adding headers for each worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total_Stock_Volume"
        
        ' Assigning headers for greatest increase and decrease of percent, ticker, and volume
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase:"
        ws.Cells(3, 15).Value = "Greatest % Decrease:"
        ws.Cells(4, 15).Value = "Greatest Total Volume:"
                            
        ' Initializing lastrow as lr for looping and assigning variable output to store from 2.
        lr = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        total = 0
        output = 2

        ' Declaring yearly and percent change
        Dim yearlychange As Double
        Dim percentchange As Double

        ' Looping on column A to check A2 and header to assign opening value and ticker
        For i = 2 To lr
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                openstock = ws.Cells(i, 3).Value
                ticker = ws.Cells(i, 1).Value
            End If
            
            ' Storing sum of total volume
            closingstock = ws.Cells(i, 6).Value
            total = total + ws.Cells(i, 7).Value
            
            ' Checking yearlychange value comparing in the loop of range A:A
            If i = lr Or ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                yearlychange = closingstock - openstock
                                        
                ' Getting percent change
                If openstock <> 0 Then
                    percentchange = (yearlychange / openstock)
                Else
                    percentchange = 0
                End If
                
                ' Putting Output values to the worksheet
                ws.Range("I" & output).Value = ticker
                ws.Range("J" & output).Value = yearlychange
                ws.Range("K" & output).Value = percentchange ' Store the numeric value of percentage
                ws.Range("K" & output).NumberFormat = "0.00%" ' Format the cell as percentage
                ws.Range("L" & output).Value = total
                
                ' Apply color formatting based on sign of yearly change
                If yearlychange < 0 Then
                    ws.Cells(output, 10).Interior.ColorIndex = 3 ' Red for negative
                Else
                    ws.Cells(output, 10).Interior.ColorIndex = 4 ' Green for positive or zero
                End If
                
                ' Moving to the next row in the output
                output = output + 1
                
                ' Reset total for the next stock
                total = 0
            End If
        Next i

        ' Find greatest increase/decrease & volume
        Dim greatestIncreaseTicker As String
        Dim greatestDecreaseTicker As String
        Dim greatestVolumeTicker As String
        Dim percentInc As Double
        Dim percentDec As Double
        Dim grtVolume As LongLong

        ' Initial values
        percentInc = 0
        percentDec = 0
        grtVolume = 0

        ' Loop through the data to find greatest increase/decrease and total volume
        ' and finding initial value greater than 0 and incrementing thereafetr with all variables.
        For i = 2 To lr
            If ws.Cells(i, 11).Value > percentInc Then
                percentInc = ws.Cells(i, 11).Value
                greatestIncreaseTicker = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11).Value < percentDec Then
                percentDec = ws.Cells(i, 11).Value
                greatestDecreaseTicker = ws.Cells(i, 9).Value
            End If

            If ws.Cells(i, 12).Value > grtVolume Then
                grtVolume = ws.Cells(i, 12).Value
                greatestVolumeTicker = ws.Cells(i, 9).Value
            End If
        Next i
            
        ' Output the greatest increase, decrease, total volume
        ws.Range("P2").Value = greatestIncreaseTicker
        ws.Range("P3").Value = greatestDecreaseTicker
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q2").Value = Format(percentInc, "percent")
        ws.Range("Q3").Value = Format(percentDec, "percent")
        ws.Range("Q4").Value = grtVolume

Next ws
End Sub

