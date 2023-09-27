Sub StockData()
    'ws.loop
    Dim ws As Worksheet
    For Each ws In Worksheets
    
    
    'Lets start small, come back for variables as you need them
        Dim Ticker As String
            Dim TickCount As Integer
            TickCount = 1
        Dim Percent As Double
            Dim first As Double
            Dim last As Double
            Dim Change As Double
        Dim Vol As Double
            Vol = 0
        Dim largecounter As Double
        Dim smallcounter As Double
        Dim volcounter As Double
            largecounter = 0
            smallcounter = 0
            volcounter = 0
            
        Dim changerange As Range
        Set changerange = ws.Range("k:l")
        changerange.FormatConditions.Delete
    
    'Add labels to the new ws.cells
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Total change"
        ws.Cells(1, 12).Value = "Percent change"
        ws.Cells(1, 13).Value = "Total volume"
            ws.Cells(1, 17).Value = "Ticker"
            ws.Cells(1, 18).Value = "Value"
            ws.Cells(2, 16).Value = "Greatest increase"
            ws.Cells(3, 16).Value = "Greatest decrease"
            ws.Cells(4, 16).Value = "Greatest volume"
                
                'readme* found thisn at https://www.automateexcel.com/vba/conditional-formatting/

                changerange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Total change"
                changerange.FormatConditions(1).Interior.ColorIndex = 2
                changerange.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="Percent change"
                changerange.FormatConditions(2).Interior.ColorIndex = 2
                
                changerange.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
                changerange.FormatConditions(3).Interior.ColorIndex = 3

                changerange.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
                changerange.FormatConditions(4).Interior.ColorIndex = 4

            
    'Find The Last Row
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'next up, the loop
        For i = 2 To lastrow
        For j = 1 To 7
        
        'calculate and print percentage
        
        If ws.Cells(i, 2).Value = 20180102 Or ws.Cells(i, 2).Value = 20190102 Or ws.Cells(i, 2) = 20200102 Then
                first = ws.Cells(i, 3).Value
            Else
            If ws.Cells(i, 2).Value = 20181231 Or ws.Cells(i, 2).Value = 20191231 Or ws.Cells(i, 2).Value = 20201231 Then
                last = ws.Cells(i, 3).Value
            End If
        End If
        
        'End the Row loop
        Next j
        
        'Calculate Volume
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Vol = 0
        End If
        Vol = Vol + ws.Cells(i, 7).Value
                
        'Ticker data
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
            Ticker = ws.Cells(i, 1).Value
            TickCount = TickCount + 1 'this code can be reused for other prints
                
            'Calculate change // percent
            Change = first - last
            Percent = (Change / last) * 100
    
                'Print data
                ws.Cells(TickCount, 10).Value = Ticker
                ws.Cells(TickCount, 11).Value = Change
                ws.Cells(TickCount, 12).Value = Percent
                ws.Cells(TickCount, 13).Value = Vol
                Range("L:L").NumberFormat = "0.00%" 'readme* we found this solution on stack overflow (https://stackoverflow.com/questions/42844778/vba-for-each-cell-in-range-format-as-percentage)
        End If
        'End the loop
        Next i
        
        
        'New loop, check for the greatests
        For i = 2 To lastrow
        For j = 10 To 13
            'find greatest increase
            If ws.Cells(i, 12).Value > largecounter Then
                largecounter = ws.Cells(i, 12).Value
                ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
                ws.Cells(2, 18).Value = largecounter
                ws.Cells(2, 18).NumberFormat = "0.00%"
            End If
            'find greatest decrease
            If ws.Cells(i, 12).Value < smallcounter Then
                smallcounter = ws.Cells(i, 12).Value
                ws.Cells(3, 17).Value = ws.Cells(i, 10).Value
                ws.Cells(3, 18).Value = smallcounter
                ws.Cells(3, 18).NumberFormat = "0.00%"
            End If
            'find greatest vol
            If ws.Cells(i, 13).Value > volcounter Then
                volcounter = ws.Cells(i, 13).Value
                ws.Cells(4, 17).Value = ws.Cells(i, 10).Value
                ws.Cells(4, 18).Value = volcounter
            End If
        Next j
        Next i
    'end the ws loop
    Next ws
    MsgBox ("DONE!")
    End Sub

