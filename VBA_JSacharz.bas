Attribute VB_Name = "Module1"
Sub Stock_Market_WallSt()

'#Set the worksheet
Dim ws As Worksheet
'There is more than 1 worksheet whcih we will be looping through
For Each ws In Worksheets

'##Add new headings for the columns:
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Vol"

'Define Ticker as String for the column "I" and also as variable to find TickerNumber=number of rows with the same Ticker till it changes to a new Ticker

Dim Ticker As String
Ticker = " "


'###Define variables used for calculation as Double, they all == 0 intitially
Dim StartYear As Double
StartYear = 0
Dim EndYear As Double
EndYear = 0
Dim PriceChange As Double
PriceChange = 0
Dim StockVol As Double
StockVol = 0
Dim PercentageChange As Double
PercentageChange = 0
Dim LastRow As Long
Dim TickerNumber As Integer
TickerNumber = 0


'####Format column "K" to set it to percetnage [%]
ws.Range("K:K").NumberFormat = "0.00%"


'Find the last row in the "A" column
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'#####Use Conditional formatting to change the cells colours accodingly to the yearly change value

'--------First condition: if J column Values = 0 then cell color will be left unchanged (white)

ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlExpression, Formula1:="=LEN(TRIM(A1))=0"
ws.Range("J2:J" & LastRow).FormatConditions(1).Interior.Pattern = xlNone


'------- Second condition: if J column Values > 0 then the cells color will be green

ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
ws.Range("J2:J" & LastRow).FormatConditions(2).Interior.ColorIndex = 4


'-------Third condition: if J column Values < 0 then the cells color will be red

ws.Range("J2:J" & LastRow).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
ws.Range("J2:J" & LastRow).FormatConditions(3).Interior.ColorIndex = 3


    
'######Introduce a new variable that will be storing the outcomes of the following calculations (as listed)
Dim NewRow As Integer
'It will start from Row =2 to skip the column name
NewRow = 2
    
   Dim i As Double
    'Loop in column "A" for the Ticker if the next row <> previous one
    'Start at position 2, skip the name of the column at (1,J)
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ws.Cells(NewRow, 9).Value = ws.Cells(i, 1).Value
    

        '#Calculate the Price Change, but first define the Start and EndYear price:
        StartYear = ws.Cells(i - TickerNumber, 3).Value
        EndYear = ws.Cells(i - TickerNumber, 6).Value
    
        PriceChange = StartYear - EndYear
    
        '##Yearly Price change should appaer in column "J"; Yearly percentage change in column "K", Total Stock Volmue in column "L":
        
        ws.Cells(NewRow, 10).Value = PriceChange
    
        ws.Cells(NewRow, 11).Value = (PriceChange / StartYear)
    
        ws.Cells(NewRow, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(i - TickerNumber, 7), ws.Cells(i, 7)))
        
        'Reset the count: TickerNumber = 0 before counting the next one
        TickerNumber = 0
        
        NewRow = NewRow + 1
        
        'If the Ticker has not changed then add +1 to the count of Ticker rows
        Else
        TickerNumber = TickerNumber + 1
                
                'End the conditional IF
                End If
                
    'Go to to the next i in the loop
   Next i
    

'Loop through the next ws
Next ws



End Sub
