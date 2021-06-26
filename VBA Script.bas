Attribute VB_Name = "Module3"
Sub Stock()


Dim worksheetname As String
Dim ws As Worksheet


Dim ticker As String
Dim tickersummary As Integer


Dim yearlychange As Double
Dim beginning As Double
Dim ending As Double

Dim percent As Double
Dim voltotal As Single
Dim vol_total As Single


For Each ws In Worksheets
ws.Activate

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"


tickersummary = 2
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

beginning = Cells(2, 3)
voltotal = Cells(2, 7)

For i = 2 To Lastrow


    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        ending = ws.Cells(i, 6).Value
        yearlychange = ending - beginning
            If beginning = 0 Then
                percent = 0
                Else: percent = yearlychange / beginning
            End If
        vol_total = voltotal + ws.Cells(i, 7).Value
        ws.Range("I" & tickersummary).Value = ticker
        ws.Range("J" & tickersummary).Value = yearlychange
        If yearlychange >= 0 Then
            ws.Range("J" & tickersummary).Interior.ColorIndex = 4
            Else
            ws.Range("J" & tickersummary).Interior.ColorIndex = 3
        End If
        
        
        ws.Range("K" & tickersummary).Value = percent
        ws.Range("K" & tickersummary).NumberFormat = "0.00%"

        ws.Range("L" & tickersummary).Value = vol_total


            
        beginning = ws.Cells(i + 1, 3).Value
        voltotal = ws.Cells(i + 1, 7).Value
        tickersummary = tickersummary + 1
        
        Else
        
        voltotal = voltotal + ws.Cells(i, 7).Value

        
    End If

Next i

ws.Columns("I:Q").EntireColumn.AutoFit




Next ws

MsgBox "Finished"

End Sub



