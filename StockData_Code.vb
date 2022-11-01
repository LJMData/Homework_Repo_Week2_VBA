Sub Stock()

'Set Vairables
Dim SD As Double

Dim PC As Double

Dim Volume As Double

Dim Start As Double

Dim Last_Row As Long
Dim i As Long

Dim Summary_Row As Integer

Dim Ticker As String

Dim ws As Worksheet

'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Main Challenge
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

'Loop Through Worskeets
For Each ws In Worksheets

    'Set Starting Values
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    SD = 0
    PD = 0
    Volume = 0
    Start = 2
    Summary_Row = 2

    'Set Loop for Current Worsksheet
    For i = 2 To Last_Row
    
            'Check Ticker Symbols
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                     
                 'Set Ticker Symbol
                 Ticker = ws.Cells(i, 1).Value
                 
                    'Add to the Ticker Totals
                    SD = ws.Cells(i, 6).Value - ws.Cells(Start, 3)
                    Volume = Volume + ws.Cells(i, 7).Value
                    PC = SD / ws.Cells(Start, 3)
                    Start = (i + 1)
                                  
                        'Print Totals
                        ws.Range("I" & Summary_Row).Value = Ticker
                        ws.Range("J" & Summary_Row).Value = SD
                        ws.Range("J" & Summary_Row).NumberFormat = "[$$-en-US]0.00"
                        ws.Range("K" & Summary_Row).Value = PC
                        ws.Range("K" & Summary_Row).NumberFormat = "0.00%"
                        ws.Range("L" & Summary_Row).Value = Volume
                
                 
                            'Conditional Fommatting for SD
                                If ws.Range("J" & Summary_Row).Value >= 0 Then
                                
                                ws.Range("J" & Summary_Row).Interior.Color = RGB(0, 255, 0)
                                
                                Else
                                
                                ws.Range("J" & Summary_Row).Interior.Color = RGB(255, 0, 0)
                                
                                End If
                                      
                 'Add a Row to the Summary Row
                 Summary_Row = Summary_Row + 1
                 
                 'Reset Totals
                 SD = 0
                 PD = 0
                 Volume = 0
                      
             Else
             
                 'Add to Totals
                 Volume = Volume + ws.Cells(i, 7).Value
                 
             End If

    Next i
       
        'Add Headers Rows
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Autofit values
        ws.UsedRange.EntireColumn.AutoFit
        
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
'Bonus
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
        
        'Calculate % increase or decrease - Set Vairables
        Dim PMax As Double
        Dim PMin As Double
        Dim r As Range
        
            'Set the Range
            Set r = ws.Range("K2:K" & ws.Rows.Count)
            
                'Find the Min and Max %
                PMax = Application.WorksheetFunction.Max(r)
                PMin = Application.WorksheetFunction.Min(r)
                
                        'Print and Format the Output
                        ws.Range("P1").Value = "Ticker"
                        ws.Range("Q1").Value = "Value"
                        ws.Range("O2").Value = "Greatest % Increase"
                        ws.Range("Q2").Value = PMax
                        ws.Range("Q2").NumberFormat = "0.00%"
                        ws.Range("O3").Value = "Greatest % Decrease"
                        ws.Range("Q3").Value = PMin
                        ws.Range("Q3").NumberFormat = "0.00%"
        
        'Calculate Greatest Volume - Set Vairables
        Dim VMax As Double
        Dim s As Range
        
            'Set the Range
            Set s = ws.Range("L2:L" & ws.Rows.Count)
        
                    'Find the Max Volume
                    VMax = Application.WorksheetFunction.Max(s)
        
                            'Pint the Output
                            ws.Range("O4").Value = "Greatest Total Volume"
                            ws.Range("Q4").Value = VMax
        
        'Insert Ticker Labels for Max %
        For i = 2 To Last_Row
        
        If ws.Cells(i, 11) = PMax Then
        
        ws.Range("P2").Value = ws.Cells(i, 9).Value
        
        Else
                
        End If
        Next i
                  
        'Insert Ticker Labels for Min %
        For i = 2 To Last_Row
        
        If ws.Cells(i, 11) = PMin Then
        
        ws.Range("P3").Value = ws.Cells(i, 9).Value
        Else
    
        End If
        Next i
              
        
        'Insert Ticker Labels for Volume
        For i = 2 To Last_Row
        
        If ws.Cells(i, 12) = VMax Then
        
        ws.Range("P4").Value = ws.Cells(i, 9).Value
        Else
       
        End If
        Next i
        
        'Autofit values
        ws.UsedRange.EntireColumn.AutoFit
              
                        
Next ws

MsgBox "Stock Analysis Complete"


End Sub
