Attribute VB_Name = "Module1"
Sub Process_all()
'
'Perform the Process_sheet procedure for each worksheet in the workbook.
'
Dim ws As Worksheet


For Each ws In Worksheets

    ws.Activate
    Process_sheet
    
Next

MsgBox "Finished"

End Sub




Sub Process_sheet()
'
'Process the data in the active worksheet
'
Dim MySheet As New Worksheet
Dim myTicker As String, PastTicker As String
Dim StartDate As Single, StopDate As Single
Dim StartPrice As Double, EndPrice As Double
Dim Pchange As Double
Dim MaxIncrease As Double, MaxDecreaseas As Double, MaxTotalVolume As Double
Dim TmaxIncrease As String, TmaxDecrease As String, TmaxTotal As String
Dim i As Single, j As Single
Dim TotalVolume As Double

'Set timer to measure the code's performance.
StartDate = VBA.DateTime.Timer


Set MySheet = ActiveSheet

With MySheet

    'Clear out the spreadsheet before writing new data and count the rows
    .Range("H:Z").Clear
    
    'Write main table headers
    .Range("I1").Value = "Ticker"
    .Range("J1").Value = "Yearly Change"
    .Range("K1").Value = "Percent Change"
    .Range("L1").Value = "Total Change"
    
    'Write summary table row lables
    .Range("O2").Value = "Greatest %Increase"
    .Range("O3").Value = "Greatest %Decrease"
    .Range("O4").Value = "Greatest Total Value"
    
    'Write summary table headers
    .Range("P1").Value = "Ticker"
    .Range("Q1").Value = "Value"
    
      
    'Count number of table rows
    RowCount = Cells(Rows.Count, 1).End(xlUp).Row
    
    j = 2
    
    'Iterate through parent table rows
    For i = 2 To RowCount
    
        'Determine the current and previous row ticker.
        myTicker = .Cells(i, 1).Value
        PastTicker = .Cells(i - 1, 1).Value
        
        'Determine if the current ticker is a new ticker.
        If myTicker <> PastTicker Then
        
            'The code within this if block is supposed to execute, if the procedure enters a new ticker,
            'while the first ticker should be an exception.
            If i > 3 Then
            
            
                EndValue = .Cells(i - 1, 6).Value
                .Cells(j - 1, 10) = EndValue - StartValue
                
                If StartValue = 0 Then
                
                    .Cells(j - 1, 11).Value = "Start-up"
                    
                
                Else
                
                    Pchange = (EndValue - StartValue) / StartValue

                    .Cells(j - 1, 11).Value = Pchange
                    
                    If Pchange > 0 Then
                    
                        .Cells(j - 1, 11).Interior.Color = vbGreen
                    
                    Else
                    
                        .Cells(j - 1, 11).Interior.Color = vbRed
                        
                    End If
                    
                
                End If
                
                .Cells(j - 1, 12) = TotalVolume
                
                'Determine if the current Total Volume is the maximum until current iteration
                If MaxTotalVolume < TotalVolume Then
                
                    MaxTotalVolume = TotalVolume
                    TmaxTotal = PastTicker
                
                End If
                
                'Determine if the current Max increase is the maximum until current iteration
                If MaxIncrease < Pchange Then
                
                    MaxIncrease = Pchange
                    TmaxIncrease = PastTicker
                
                End If
                
                'Determine if the current Max increase is the maximum until current iteration
                If MaxDecrease > Pchange Then
                
                    MaxDecrease = Pchange
                    TmaxDecrease = PastTicker
                
                End If
                
                
            
            End If
        
        
                .Cells(j, 9).Value = myTicker
                StartValue = .Cells(i, 3).Value
                TotalVolume = .Cells(i, 7).Value
            
                j = j + 1
        
        Else
        
            TotalVolume = TotalVolume + .Cells(i, 7).Value
            
        
        End If
    
 
    
    Next i

    .Range("Q2").Value = MaxIncrease
    .Range("Q3").Value = MaxDecrease
    .Range("Q4").Value = MaxTotalVolume
    
    .Range("P2").Value = TmaxIncrease
    .Range("P3").Value = TmaxDecrease
    .Range("P4").Value = TmaxTotal



    .Columns.AutoFit
    .Columns("K:K").NumberFormat = "0.00%"
    .Range("Q2:Q3").NumberFormat = "0.00%"
    .Range("Q4").NumberFormat = "0.0000E+00"
    
10    StopDate = VBA.DateTime.Timer

    .Range("O16").Value = "Process duration (s)"
    .Range("P16").Value = Round(StopDate - StartDate, 5)


End With


End Sub


