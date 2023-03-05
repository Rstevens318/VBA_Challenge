Attribute VB_Name = "Module2"
Sub singleWKsheets()

        
'Set Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
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
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        FDValue = Cells(2, 3).Value
        Range("K2:K" & LastRow).NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        Dim i As LongLong
      
'Start Loop

            For i = 2 To LastRow

'Set Variables within loop

                PrevName = Cells(i + 1, 1).Value
                Name = Cells(i, 1).Value
                LDValue = Cells(i, 6).Value
                
'Start Conditions If this does not equal this
                
                If PrevName <> Name Then
                               
'Set YearlyChange

                YearOpen = FDValue
                 YearChange = LDValue - YearOpen
                  Cells(SummaryRow, 10).Value = YearChange
                  
'Set YearChange Color
                
                    If YearChange <= 0 Then
                     Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    Else
                     Cells(SummaryRow, 10).Interior.ColorIndex = 4
                    End If
                     
'Set Percent Change

                   PercentChange = (LDValue - YearOpen) / YearOpen
                    Cells(SummaryRow, 11).Value = PercentChange
                    
'Set Ticker Name and Populate

                 Ticker = Name
                  Cells(SummaryRow, 9).Value = Ticker

'Set Total Volume per Ticker and populate

                   TotalVolume = TotalVolume + Cells(i, 7).Value
                    Cells(SummaryRow, 12).Value = TotalVolume
                    
'Set Next Row Values
                  
                     SummaryRow = SummaryRow + 1
                    TotalVolume = 0
                   YearChange = 0
                  PercentageChange = 0
                 FDValue = Cells(i + 1, 3).Value
                   
'Conditions if ticker names are equal

                 Else
                 
                        TotalVolume = TotalVolume + Cells(i, 7).Value
                         
                         
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
            
             Set Percentrng = Range("K2:K" & Rows.Count)
             Set VolumeRng = Range("L2:L" & Rows.Count)
             
'Set Min and Max and Total Volume

            min = Application.WorksheetFunction.min(Percentrng)
             max = Application.WorksheetFunction.max(Percentrng)
              MaxVolume = Application.WorksheetFunction.max(VolumeRng)
              
                Cells(2, 17).Value = max
                Cells(3, 17).Value = min
                Cells(4, 17).Value = MaxVolume
           
'Loop For Ticker Name

            For j = 2 To LastRow
            
            Name = Cells(j, 9).Value
             If max = Cells(j, 11).Value Then
              Cells(2, 16).Value = Name
              
             ElseIf Cells(j, 11).Value = min Then
              Cells(3, 16) = Name
              
            ElseIf Cells(j, 12).Value = MaxVolume Then
             Cells(4, 16).Value = Name
             
            End If
            
            Next j
            

End Sub

