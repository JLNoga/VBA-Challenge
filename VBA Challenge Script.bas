Attribute VB_Name = "Module1"
'Add new columns for:
'Ticker | Yearly Change | Percent Change | Total Stock Volume
'Loop through each stock to find the above data and input in new columns
'Continue for each tab

'Inserts new columns in I(9) J(10) K(11) L(12)
'Ticker | Yearly Change | Percent Change | Total Stock Volume

'Completes for each worksheet

Sub NewColumns()
    Dim currentsheet As Worksheet
    For Each currentsheet In Worksheets
        currentsheet.Range("I1").EntireColumn.Insert
        currentsheet.Range("I1").Value = "Total Stock Volume"
        currentsheet.Range("I1").EntireColumn.Insert
        currentsheet.Range("I1").Value = "Percent Change"
        currentsheet.Range("I1").EntireColumn.Insert
        currentsheet.Range("I1").Value = "Yearly Change"
        currentsheet.Range("I1").EntireColumn.Insert
        currentsheet.Range("I1").Value = "Ticker"
    Next currentsheet
End Sub

'Finds change in ticker name (assumes 1-year for each)
'Adds ticker name to Ticker column (I)
'calculates and adds yearly change to column J
'calculates and adds percentage change to column K
'calculates and adds total volume to column L

Sub TickerChange()
    Dim currentsheet As Worksheet
    Dim ticker As String
    Dim yearly_change As Double
    Dim percent_change As Double
    'Define total volume as LongPtr to stop overflow error - this macro will only work on 64 bit
    'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/longptr-data-type
    Dim total_volume As LongPtr
    Dim summary_table_row As Integer
    Dim LR As Long
    Dim PriceOpen As Double
    Dim PriceClose As Double
    Dim Row As Long
    'set beginning variable values
    total_volume = 0
    summary_table_row = 2
    
    
    
    
    For Each currentsheet In Worksheets
    PriceOpen = currentsheet.Cells(2, 3).Value
    PriceClose = 0
    LR = currentsheet.Cells(Rows.Count, 1).End(xlUp).Row
    
        For Row = 2 To LR
                       
            If currentsheet.Cells(Row + 1, 1).Value <> currentsheet.Cells(Row, 1).Value Then
                
                'calculate the variables
                ticker = currentsheet.Cells(Row, 1).Value
                total_volume = total_volume + currentsheet.Cells(Row, 7).Value
                PriceClose = currentsheet.Cells(Row, 6).Value
                yearly_change = PriceClose - PriceOpen
                percent_change = yearly_change / PriceOpen
                
                'set the values & format in the summary table
                currentsheet.Range("I" & summary_table_row).Value = ticker
                currentsheet.Range("J" & summary_table_row).Value = yearly_change
                If currentsheet.Range("J" & summary_table_row).Value >= 0 Then
                    currentsheet.Range("J" & summary_table_row).Interior.ColorIndex = 4
                Else
                    currentsheet.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                currentsheet.Range("K" & summary_table_row).Value = percent_change
                currentsheet.Range("K" & summary_table_row).NumberFormat = "0.0%"
                currentsheet.Range("L" & summary_table_row).Value = total_volume
                
                'reset the summary table and variables
                summary_table_row = summary_table_row + 1
                PriceClose = 0
                total_volume = 0
                
                'set the open price to the value of the next ticker open
                    If currentsheet.Cells(Row + 1, 3).Value > 0 Then
                        PriceOpen = currentsheet.Cells(Row + 1, 3).Value
                    Else
                        PriceOpen = currentsheet.Cells(2, 3).Row
                        summary_table_row = 2
                    End If
            Else
                'adds the current volume to the total volume
                total_volume = total_volume + currentsheet.Cells(Row, 7).Value
            End If
        Next Row
    Next currentsheet
End Sub

Sub greatestTable()
'Create Summary table with greatest % increase, % decrease, and total volume

Dim currentsheet As Worksheet
'Dim greatIncr As Double
'Dim greatDecr As Double
'Dim greatVol As LongPtr
Dim greatIncrTick As String
Dim greatDecrTick As String
Dim greatVolTick As String
Dim Maximum_range As Range
Dim Minimum_range As Range
Dim Maximum_range_two As Range
Dim LRSum As Long


    For Each currentsheet In Worksheets
    'MsgBox (greatVolTick)
    LRSum = currentsheet.Cells(Rows.Count, 10).End(xlUp).Row
    Set Maximum_range = currentsheet.Range("K2", "K" & LRSum)
    Set Minimum_range = currentsheet.Range("K2", "K" & LRSum)
    Set Maximum_range_two = currentsheet.Range("L2", "L" & LRSum)
    'MsgBox (Str(Maximum_range_two.Count))
    

        'create columns & titles for summary table
        currentsheet.Range("P1").EntireColumn.Insert
        currentsheet.Range("P1").Value = "Value"
        currentsheet.Range("P1").EntireColumn.Insert
        currentsheet.Range("P1").Value = "Ticker"
        currentsheet.Range("O2").Value = "Greatest % Increase"
        currentsheet.Range("O3").Value = "Greatest % Decrease"
        currentsheet.Range("O4").Value = "Greatest Total Volume"
        
        'set values for summary table variables
        'max increase
        
        For Each cell In Maximum_range
            If cell.Value > i Then
                i = cell.Value
                j = cell.Offset(0, -2).Value
            End If
        Next
        greatIncr = i
        greatIncrTick = j
                
        'Min increase
        
        For Each cell In Minimum_range
            If cell.Value < k Then
                k = cell.Value
                l = cell.Offset(0, -2).Value
            End If
        Next
        greatDecr = k
        greatDecrTick = l
        
        'Max volume
        
        For Each cell In Maximum_range_two
            If cell.Value > m Then
                m = cell.Value
                n = cell.Offset(0, -3).Value
            End If
        Next
        greatVol = m
        greatVolTick = n
        
        'assign values to summary table
        currentsheet.Range("Q2").Value = greatIncr
        currentsheet.Range("Q3").Value = greatDecr
        currentsheet.Range("Q2:Q3").NumberFormat = "0.0%"
        currentsheet.Range("Q4").Value = greatVol
        currentsheet.Range("Q4").NumberFormat = "0"
        
        currentsheet.Range("P2").Value = greatIncrTick
        currentsheet.Range("P3").Value = greatDecrTick
        currentsheet.Range("P4").Value = greatVolTick
        
        i = 0
        j = ""
        k = 0
        l = ""
        m = 0
        n = ""
        
        
    Next currentsheet
End Sub


