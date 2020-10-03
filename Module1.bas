Attribute VB_Name = "Module1"
Sub homework()
D = Worksheets.Count
MsgBox (D)
'I changed close - end and added D, changed conditional formatting and volume is wrong add +1, add extra credit, end For loop at 705715


'FIND BEGGINIG AND END OF TICKERS

For W = 1 To D 'look through all worksheets
Worksheets(W).Activate
Cells(1, 9) = "Ticker" 'Title
Cells(2, 9) = Cells(2, 1)
j = 2
  For i = 2 To 705715
     Column = 1
    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    Cells(j, 9) = Cells(i, Column).Value
    j = j + 1
    End If
  Next i
Next W





'FIND CHANGE (OPEN BEG - CLOSE END)
For W = 1 To D
Worksheets(W).Activate
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"    'TITLES
Cells(1, 13) = "Closing at End"
Cells(1, 14) = "Opening at Beg"
j = 1

  For i = 2 To 705715
     Column = 1
    ' Searches for when the value of the next cell is different than that of the current cell
    If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
    Cells(2, 14) = Cells(2, 3).Value
    Cells(j + 1, 13) = Cells(i, 6).Value 'end close VISUAL
    Cells(j + 2, 14) = Cells(i + 1, 3).Value 'open beg
    Cells(j + 1, 10) = -Cells(j + 1, 14).Value + Cells(j + 1, 13).Value 'DIFFERENCE OPEN BEG - CLOSE END YEARLY CHANGE
     
         'CASE: when denominator is zero
         If Cells(j + 1, 14) = 0 Then
            Cells(j + 1, 14) = 1
         End If
     
    Cells(j + 1, 11) = FormatPercent(Cells(j + 1, 10) / Cells(j + 1, 14)) 'PERCENT CHANGE
    Cells(j + 1, 16) = i  'location of end of tickers
    
        'Conditional Formatting
         If Cells(j + 1, 10) > 0 Then
          Cells(j + 1, 10).Interior.ColorIndex = 4 'Positive Green
          Else
          Cells(j + 1, 10).Interior.ColorIndex = 3 'Negative Red
          End If
          
     'first sum for volume
    Dim Total1 As Double
    Total1 = WorksheetFunction.Sum(Range(Cells(1, 7), Cells(Cells(2, 16), 7)))
    Cells(2, 12) = Total1
  
   j = j + 1 'step thru

    End If
  Next i
Next W


'TOTAL VOLUME LOOP
For W = 1 To D
Worksheets(W).Activate
n = WorksheetFunction.CountA(Columns(16)) 'length of each ticker to sum
r = 3
    For V = 2 To n
    begin = Cells(V, 16) + 1
    last = Cells(V + 1, 16)
    Dim Total As Double
    Total = WorksheetFunction.Sum(Range(Cells(begin, 7), Cells(last, 7)))
    Cells(r, 12) = Total
    r = r + 1
    Next V
Next W



For W = 1 To D
Worksheets(W).Activate
Worksheets(W).Columns(13).ClearContents
Worksheets(W).Columns(14).ClearContents  'erase the columns i used for visualizing
Worksheets(W).Columns(16).ClearContents
MaxValue = Application.WorksheetFunction.Max(Columns(12))
MaxValue2 = FormatPercent(Application.WorksheetFunction.Max(Columns(11)))
MaxValue3 = FormatPercent(Application.WorksheetFunction.Min(Columns(11)))
Cells(4, 19) = MaxValue
Cells(2, 19) = MaxValue2
Cells(3, 19) = MaxValue3
Cells(1, 18) = "Ticker"
Cells(1, 19) = "Value"
Cells(2, 17) = "greatest % increase"
Cells(3, 17) = "greatest % decrease"
Cells(4, 17) = "greatest total volume"

For i = 2 To 705715

If Cells(i, 12) = MaxValue Then
Cells(4, 18) = Cells(i, 9)
End If

If Cells(i, 11) = Cells(2, 19) Then
Cells(2, 18) = Cells(i, 9)
End If

If Cells(i, 11) = Cells(3, 19) Then
Cells(3, 18) = Cells(i, 9)
End If


Next i


Next W







End Sub
