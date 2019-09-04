Attribute VB_Name = "Module1"
Sub Calc_StockInfoBySheet()

    
    
    
    Dim EndVal As Double        'Last Closing Value
    Dim GIncrease As Double     'Holder for Greatest Increase
    Dim GDecrease As Double     'Holder for Greatest Decrease
    Dim GVolume As Double       'Holder for Greatest Volume
    Dim GTicker(3) As String    'Holder for ticker symbol
    Dim LastRow As Long
    Dim starting_ws As Worksheet
    Dim StartVal As Double      'First Closing Value
    Dim SumRowNum As Integer    'Summary Info Row Number
    Dim TickerCnt As Integer    'Indicator to set starting value for %Change calculations
    Dim Volume As Double        'Total Volume
    Dim ws_num As Integer
   
      

    
    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
    ws_num = ThisWorkbook.Worksheets.Count

   

    For shNum = 1 To ws_num
        ThisWorkbook.Worksheets(shNum).Activate
        
        'find last row on active sheet
        With ActiveSheet
            LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
        ThisWorkbook.Worksheets(shNum).Cells(1, 9).Value = "Ticker"
        ThisWorkbook.Worksheets(shNum).Cells(1, 10).Value = "Yearly Change"
        ThisWorkbook.Worksheets(shNum).Cells(1, 11).Value = "Percent Change"
        ThisWorkbook.Worksheets(shNum).Cells(1, 12).Value = "Total Stock Volume"
        ThisWorkbook.Worksheets(shNum).Cells(1, 15).Value = "Ticker"
        ThisWorkbook.Worksheets(shNum).Cells(1, 16).Value = "Value"
                
        ThisWorkbook.Worksheets(shNum).Cells(2, 14).Value = ThisWorkbook.Worksheets(shNum).Name + " Greatest %Increase"
        ThisWorkbook.Worksheets(shNum).Cells(3, 14).Value = ThisWorkbook.Worksheets(shNum).Name + " Greatest %Decrease"
        ThisWorkbook.Worksheets(shNum).Cells(4, 14).Value = ThisWorkbook.Worksheets(shNum).Name + " Greatest Total Volume"
        
        
         'Set counter for summary rows
        SumRowNum = 2
    
        'Set holders for Greatest values
        GIncrease = 0
        GDecrease = 0
        GVolume = 0
        
        
        TickerCnt = 1
        Volume = 0
                
        For rNum = 2 To LastRow
            If TickerCnt = 1 Then
                StartVal = Cells(rNum, 6).Value
            End If
            
            Volume = Volume + Cells(rNum, 7).Value
            
            'Determine if ticker changes next row, if so set summary values
            If Cells(rNum, 1).Value <> Cells(rNum + 1, 1).Value Then
                
                'Set Ending Value to calc yearly/% change
                EndVal = Cells(rNum, 6).Value
                
                'Set summary info - Ticker, Yearly Change, %Change, Total Volume
                ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 9).Value = Cells(rNum, 1).Value 'Ticker
                ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 10).Value = EndVal - StartVal 'Yearly Change
                
                'Set Color based on Yearly Change
                If EndVal - StartVal >= 0 Then
                    ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 10).Interior.ColorIndex = 4
                                     
                Else
                    ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 10).Interior.ColorIndex = 3
                 
                End If
                    
                'Percent Change
                If EndVal > 0 And StartVal > 0 Then 'test for zero's (PLNT)
                    ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value = (EndVal - StartVal) / StartVal * 100
                    
                    'Test for greatest increase/decrease
                    If ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value > 0 Then
                        If ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value > GIncrease Then
                            GIncrease = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value
                            GTicker(1) = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 9).Value
                        End If
                    Else 'decrease
                        If ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value < GDecrease Then
                            GDecrease = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 11).Value
                            GTicker(2) = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 9).Value
                        End If
                    End If
                    
                End If
                                           
                'Total Volume
                ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 12).Value = Volume
                
                'Test for greatest volume
                If ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 12).Value > GVolume Then
                    GVolume = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 12).Value
                    GTicker(3) = ThisWorkbook.Worksheets(shNum).Cells(SumRowNum, 9).Value
                End If
                
                
                
                
                'Reset variables for next ticker symbol
                TickerCnt = 1
                Volume = 0
                SumRowNum = SumRowNum + 1
            Else
                TickerCnt = TickerCnt + 1
            End If
        Next rNum
      
        ThisWorkbook.Worksheets(shNum).Cells(2, 15).Value = GTicker(1)
        ThisWorkbook.Worksheets(shNum).Cells(2, 16).Value = GIncrease
        ThisWorkbook.Worksheets(shNum).Cells(3, 15).Value = GTicker(2)
        ThisWorkbook.Worksheets(shNum).Cells(3, 16).Value = GDecrease
        ThisWorkbook.Worksheets(shNum).Cells(4, 15).Value = GTicker(3)
        ThisWorkbook.Worksheets(shNum).Cells(4, 16).Value = GVolume
     
    Next shNum

   starting_ws.Activate 'activate the worksheet that was originally active

End Sub

Sub ClearValuesAllSheets()

    Dim ws_num As Integer
    
    Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning
    ws_num = ThisWorkbook.Worksheets.Count

    For shNum = 1 To ws_num
        ThisWorkbook.Worksheets(shNum).Range("I:P").Interior.ColorIndex = 0
        ThisWorkbook.Worksheets(shNum).Range("I:P").Value = ""
    Next shNum
    
End Sub

