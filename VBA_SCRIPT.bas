Attribute VB_Name = "Module1"

'Homework-2 / VBAStocks
Sub Alphabetical_testing_Analysis()

For Each ws In Worksheets '<------  opening  worksheet loop

' column Header /Data field label
ws.Range("I1").Value = "tiker"
ws.Range("J1").Value = "yearly change from opening price"
ws.Range("K1").Value = "percentage of change"
ws.Range("P1").Value = "tiker"
ws.Range("Q1").Value = "Value"
ws.Range("L1").Value = " Total stock volume"
ws.Range("O2").Value = "Greatest % increasing"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"

Dim i As Long
Dim tiker As String
Dim targetrow As Integer
Dim row As Long
Dim lastrow As Long
Dim totalvolume As Double
Dim yearlyopen As Double
Dim yearlyclose As Double
Dim YearlyChange As Double
Dim Vcount As Long
Dim Percentage_Change As Double


Vcount = 2
totalvolume = 0
targetrow = 2

'Finding LAstRow
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).row

For row = 2 To lastrow        ' <---- opening first loop

totalvolume = totalvolume + ws.Cells(row, 7).Value

If ws.Cells(row, 1).Value <> ws.Cells(row + 1, 1).Value Then
ws.Range("I" & targetrow).Value = ws.Cells(row, 1).Value
ws.Range("L" & targetrow).Value = totalvolume

totalvolume = 0

yearlyclose = ws.Range("F" & row).Value
yearlyopen = ws.Range("C" & Vcount).Value
YearlyChange = yearlyclose - yearlyopen
ws.Range("J" & targetrow).Value = YearlyChange
 
 '% change from opening price at the beginning of a given year to the closing price at the end of that year

If yearlyclose = 0 And yearlyopen = 0 Then
Percentage_Change = 0
Else

' % of change formula

Percentage_Change = YearlyChange / yearlyopen
ws.Range("k" & targetrow) = Percentage_Change

End If
'Formating Cells to %
ws.Range("K" & targetrow).NumberFormat = "0.00%"


targetrow = targetrow + 1
  Vrow = row + 1
    
    End If
        
        'conditional formating / highlight positive change in green and negative change in red
        
            If ws.Range("J" & targetrow).Value < 0 Then

ws.Range("J" & targetrow).Interior.ColorIndex = 22

            Else

ws.Range("J" & targetrow).Interior.ColorIndex = 43

            End If

    Next row            ' <-----  closing first loop
         
 lastrow = ws.Cells(Rows.Count, 11).End(xlUp).row
 
 ' Condition to retrieve  the "Greatest % increase", "Greatest % decrease" and "Greatest total volume"
         
                 For i = 2 To lastrow      ' <------ opening second loop
                 
        If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                 
                 ws.Range("P2").Value = ws.Range("I" & i).Value
                 ws.Range("Q2").Value = ws.Range("K" & i).Value
                 
            End If
        
        If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                 
                 ws.Range("P3").Value = ws.Range("I" & i).Value
                 ws.Range("Q3").Value = ws.Range("K" & i).Value
            End If
        
        If ws.Range("K" & i).Value > ws.Range("Q4").Value Then
                 
                 ws.Range("P4").Value = ws.Range("I" & i).Value
                 ws.Range("Q4").Value = ws.Range("K" & i).Value
    
            End If

         Next i                 '<------ closing second loop
         
'Formating Cells to %

ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"

'format column to autofit  /Goal: set the column with width length based on data length

ws.Columns("I:Q").AutoFit
  
    Next ws  '<---- closing worksheet loop
 
 End Sub


