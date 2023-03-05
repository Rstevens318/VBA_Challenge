Attribute VB_Name = "Module1"
Sub multipleWKsheets()

    
'Run SIMPLE Conditionals For Worksheets
        For Each ws In Worksheets
        
'Set Headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
'Define Variables and Contstants
        Dim TotalVolume As LongLong
        Dim Ticker As String
        Dim YearChange As Double
        Dim PercentChange As Double
        Dim SummaryRow As Integer
        Dim FDValue As Double
        Dim LDValue As Double
        Dim YearOpen As Double
        Dim YearClose As Double
        Dim rng As Range
        TotalVolume = 0
        YearChange = 0
        PercentChange = 0
        SummaryRow = 2
    
'Define Starting Point for opening value and  Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        FDValue = ws.Cells(2, 3).Value
        ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        Dim i As LongLong
      
'Start Loop

            For i = 2 To LastRow

'Set Variables within loop

                PrevName = ws.Cells(i + 1, 1).Value
                Name = ws.Cells(i, 1).Value
                LDValue = ws.Cells(i, 6).Value
                
'Start Conditions If this does not equal this
                
                If PrevName <> Name Then
                               
'Set YearlyChange

                YearOpen = FDValue
                 YearChange = YearOpen - LDValue
                  ws.Cells(SummaryRow, 10).Value = YearChange
                  
'Set YearChange Color
                
                    If YearChange >= 0 Then
                     ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    Else
                     ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    End If
                     
'Set Percent Change

                   PercentChange = (LDValue - YearOpen) / YearOpen
                    ws.Cells(SummaryRow, 11).Value = PercentChange
                    
'Set Ticker Name and Populate

                 Ticker = Name
                  ws.Cells(SummaryRow, 9).Value = Ticker

'Set Total Volume per Ticker and populate

                   TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                    ws.Cells(SummaryRow, 12).Value = TotalVolume
                    
'Set Next Row Values
                  
                     SummaryRow = SummaryRow + 1
                    TotalVolume = 0
                   YearChange = 0
                  PercentageChange = 0
                 FDValue = ws.Cells(i + 1, 3).Value
                   
'Conditions if ticker names are equal

                 Else
                 
                        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                         
                         
'End Conditional

                End If
         
'End Loop

            Next i

'Find Greatest Increase, Decrease, and Total Volume
            
            Dim min As Double
            Dim max As Double
            Dim MaxVolume As LongLong
            Dim Percentrng As Range
            Dim VolumeRng As Range
            
             Set Percentrng = ws.Range("K2:K" & Rows.Count)
             Set VolumeRng = ws.Range("L2:L" & Rows.Count)
             
'Set Min and Max and Total Volume

            min = ws.Application.WorksheetFunction.min(Percentrng)
             max = ws.Application.WorksheetFunction.max(Percentrng)
              MaxVolume = ws.Application.WorksheetFunction.max(VolumeRng)
              
                ws.Cells(2, 17).Value = max
                ws.Cells(3, 17).Value = min
                ws.Cells(4, 17).Value = MaxVolume
           
'Loop For Ticker Name

            For j = 2 To LastRow
            
            Name = ws.Cells(j, 9).Value
             If max = ws.Cells(j, 11).Value Then
              ws.Cells(2, 16).Value = Name
              
             ElseIf ws.Cells(j, 11).Value = min Then
              ws.Cells(3, 16) = Name
              
            ElseIf ws.Cells(j, 12).Value = MaxVolume Then
             ws.Cells(4, 16).Value = Name
             
            End If
            
            Next j
            
                
'End Worksheet Loop
       
        Next ws


End Sub

