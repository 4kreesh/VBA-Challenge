Stock Analysis
'*************************
'       VBA Challenge 
'**************************

Sub Stock_Analysis():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim i As Long
        Dim j As Long
        j = 2
        Dim LRowA As Long
        Dim LRowI As Long
        Dim TCount As Long
        TCount = 2
        Dim PerChange As Double
        Dim GreatIncr As Double
        Dim GreatDecr As Double
        Dim GreatVol As Double

        WorksheetName = ws.Name
        
        LRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (LRowA)
        
        ws.Cells(1, 9).Value = "Ticker Symbol"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
   '******************************************************************************************************************
   'PART 1 - Ticker symbol,yearly change, color formatting, percentage change, total stock volume
   '******************************************************************************************************************
  
            For i = 2 To LRowA
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then  '----- findingnext ticker symbol when it's changed
                    ws.Cells(TCount, 9).Value = ws.Cells(i, 1).Value  '----- copying the ticker symbol to new cell I
                    ws.Cells(TCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value '----- Yearly chNge
                
                    If ws.Cells(TCount, 10).Value < 0 Then        ' -------- color formatting if the yearly change is < 0 -red, else gree
                        ws.Cells(TCount, 10).Interior.ColorIndex = 3
                    Else
                        ws.Cells(TCount, 10).Interior.ColorIndex = 4
                    End If
                    
                    If ws.Cells(j, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value) '----- Percentage change calculation
                         ws.Cells(TCount, 11).Value = Format(PerChange, "Percent")
                    Else
                          ws.Cells(TCount, 11).Value = Format(0, "Percent")
                    End If
                    
                    If ws.Cells(TCount, 11).Value < 0 Then        ' -------- color formatting if the percentage change is < 0 -red, else green
                        ws.Cells(TCount, 11).Interior.ColorIndex = 3
                    Else
                        ws.Cells(TCount, 11).Interior.ColorIndex = 4
                    End If
                    
                ws.Cells(TCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))  '------ Total Volume calculation
                TCount = TCount + 1
                j = i + 1
                
                End If
            
            Next i
            
    
'**********************************************************************
' PART 2 Greatest Volume, Greatest % Increase, Greatest % decrease
'**********************************************************************
        
        LRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        'MsgBox ( LRowI)

        GreatVol = ws.Cells(2, 12).Value '------------ Assigning cell 11 and 12 values to new variable for easy handling
        GreatIncr = ws.Cells(2, 11).Value
        GreatDecr = ws.Cells(2, 11).Value
        
            
            For i = 2 To LRowI
            
                If ws.Cells(i, 12).Value > GreatVol Then  '----- checking for next biggest value, if found assigning to greatest volumr cell
                    GreatVol = ws.Cells(i, 12).Value
                    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
               Else
                    GreatVol = GreatVol
                End If
                
                If ws.Cells(i, 11).Value > GreatIncr Then  '---- checking for next big value, if found assigning to greatest increment cell
                    GreatIncr = ws.Cells(i, 11).Value
                    ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                Else
                     GreatIncr = GreatIncr
               End If
                
                
                If ws.Cells(i, 11).Value < GreatDecr Then   '------ checking for smallest value, if found assigning it to greatest decrease cell
                    GreatDecr = ws.Cells(i, 11).Value
                    ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                Else
                    GreatDecr = GreatDecr
                End If
                
                                ws.Cells(2, 17).Value = Format(GreatIncr, "Percent")
                                ws.Cells(3, 17).Value = Format(GreatDecr, "Percent")
                                ws.Cells(4, 17).Value = Format(GreatVol, "Scientific")
            
            Next i
    
    Next ws
        
End Sub





References
Basic VBA learning https://www.homeandlearn.org/for_each.html
apply changes to all sheets in workbook
https://learn.microsoft.com/en-us/office/vba/api/excel.range.rows#syntax
worksheet functions 
https://learn.microsoft.com/en-us/office/vba/excel/concepts/events-worksheetfunctions-shapes/using-excel-worksheet-functions-in-visual-basic
