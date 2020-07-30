Attribute VB_Name = "Module11"
Sub main()

Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate


Dim ticker As String
Dim tickerVolume As Double

Dim yearlyChange As Double

Dim percentChange As Double

Dim tableRow As Integer
Dim openPrice As Double
Dim closePrice As Double
Dim priceDate As Long
Dim LastRow As Long

tableRow = 2

' adding title to cells
Range("I1").value = "Ticker"
Range("J1").value = "Yearly Change"
Range("K1").value = "Percent Change"
Range("L1").value = "Total Stock Volume"



'Loop'
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
yearOpen = Cells(2, 3).value
percentChange = 0
yearlyChange = 0
tickerVolume = 0
For i = 2 To LastRow

    'Here we are looping to get ticker
    If Cells(i + 1, 1).value <> Cells(i, 1).value Then
    Cells(tableRow, 9).value = Cells(i, 1).value
    
    'Calculate yearly change
    yearClose = Cells(i, 6).value
    yearlyChange = yearClose - yearOpen
    
    'Putting results in the right column
    Cells(tableRow, 10).value = yearlyChange
    

    
    'Calculate percent change
        If (yearClose = 0 And yearOpen = 0) Then
        percentChange = 0
        ElseIf (yearClose <> 0 And yearOpen = 0) Then
        percentChange = 1
        Else: percentChange = (yearClose - yearOpen) / yearOpen
        'Put results in the right column
        Cells(tableRow, 11).value = percentChange
        Cells(tableRow, 11).NumberFormat = "0.00%"
        
        End If
    
    'Adding ticker volume
    tickerVolume = tickerVolume + Cells(i, 7).value
    Cells(tableRow, 12).value = tickerVolume
        
    tableRow = tableRow + 1
    yearlyChange = 0
    percentChange = 0
    tickerVolume = 0
    
    Else: tickerVolume = tickerVolume + Cells(i, 7).value
   
    End If
Next i


'Coloring

Dim yearlyColor As Long
YearlyRow = Cells(Rows.Count, 10).End(xlUp).Row

For j = 2 To YearlyRow
    If Cells(j, 10).value < 0 Then
    Cells(j, 10).Interior.ColorIndex = 3
    Else: Cells(j, 10).Interior.ColorIndex = 4
    End If
    
Next j

ws.Cells(1, 1) = "<ticker>" 'this sets cell A1 of each sheet to "1"
Next

starting_ws.Activate 'activate the worksheet that was originally active



End Sub



