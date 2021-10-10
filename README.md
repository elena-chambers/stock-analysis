
# Refactoring VBA Script to Measure Stock and Code Performance.
  
  
## Overview of Project
A client with experience in the finance industry has requested assistance to create a script that will compute Total Daily Volumes and Return Percentanges for twelve green energy stocks in 2017 and 2018. The clients are espcecially interested in investing in DAQO New Energy Corp (ticker; "DQ"), however they would like to compare DQ returns to other green energy stock before commiting to DQ. This script will summarize the stock data such that they will be better equipped to advise their clients into which green energy stocks to invest. 
Since the stock data was stored in Excel sheets, the script was written in Visual Basic for Applications (VBA) so that it could be seamlessly integrated into the exisiting workspace.

### Purpose
 Assist the client develop a better understanding of green energy stock trends to advise future investors. 
 
## Results
The performace of the twelve stocks in 2017 and 2018 are compared below: 
<p float="left">
  <img src="https://user-images.githubusercontent.com/91163155/136715142-a6daee92-a252-4bd5-8496-34b64a0cdfc5.png" width="220" />
  <img src="https://user-images.githubusercontent.com/91163155/136715139-3d75c8dc-be61-4f27-ab1e-4e889845f232.png" width="220" /> 
</p>

From these tables, we can see that while DQ had the highest return (199.4%) in 2017, it had the lowest return (-62.6%) in 2018. The sharp decline in returns year-to-year might imply instablity in DQ, and other more stable stocks (like ENPH) may be more financially safe. 

### Inital & Refactored Code
The intial code used one dimensional variables that required the loop to repeat for each ticker. While this produced accurate results, it was considerably slower than the refactored code. In the refactored code, data was stored in multidimesional varbibles that allowed the code to run once through. 
##### Speed of Original Code
<p float="left">
  <img src="https://user-images.githubusercontent.com/91163155/136715636-234fad68-70a0-4f96-867c-320055b4c016.png" width="220" />
  <img src="https://user-images.githubusercontent.com/91163155/136715639-e4a925ff-26c9-4ddc-964e-ab92f17cda75.png" width="220" /> 
</p>

##### Speed of Refactored Code

<p float="left">
  <img src="https://user-images.githubusercontent.com/91163155/136715670-cf6fdaf8-5c38-4373-94f3-7631397c9ad9.png" width="220" />
  <img src="https://user-images.githubusercontent.com/91163155/136715674-6206dd7f-0239-418f-8ff8-991e1e859eba.png" width="220" /> 
</p>

The refactored code is 432% faster for the 2017 data set, and 442% faster for the 2018 set. <br/>
Below is the refactored code:

    Sub AllStocksAnalysisRefactored()
    ' Define variables that will store code timer data
    Dim startTime As Single
    Dim endTime  As Single
    
    'Create message box to allow user to specify year
    yearValue = InputBox("What year would you like to run the analysis on?")

    'Begin the code timer
    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet; create title and column headers
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String 'the (12) specifies the dimensions of the tickers array -- 12 tickers -> 12 dimesions
        'Instructions assigned tickers as pString type, but the code wouldn't run with "pString", changed to regular String
        
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
        'Array dimesions hard coded for 12 tickers, can use empty brackets () for dynamic arrays: https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-arrays
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        ticker = tickerIndex
        tickerVolumes(tickerIndex) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value '<-this line is from the Hints section
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
              
        End If
        
        '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            '3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
             
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed          
        End If    
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub




<br/>

## Summary
#### Advantages and Disadvantages of Refactoring Code
The purpose of refactoring is to streamline the program while maintaining the original functionality of the code. Refactoring can greatly reduce computational costs associated with increases in dataset volumes or changes in inputs (in this case, adding more years or different stocks). The disadvantages of refactoring include the time investment needed to refactor. In this project, refactoring the code took longer than writing the intial code. However, now that the code as been *moderatly* optimized, it will be improve its usability and performace.
<br/>
Additonal refactoring could be neccesary; the number of tickers and the ticker ids are hardcoded into the program. This could pose major issues in that the analysis is limited to these specific stocks. If the client needs this program to analyze other stocks, more work would be needed. However, this would incur additioanl cost to the client, and may prove to be a major drawback of contining the refactoring of the code.
