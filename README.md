# Stock-Analysis
CLick here to view the Excel File: [VBA Challenge](https://github.com/paree1kd/Stock-Analysis/blob/ebc1a7ada512698e6a3a6d7dea000db5a5196c83/VBA_Challenge.xlsm)

**Overview of Project:** 
The purpose of this project was to improve "refactor" the VBA code provided that showed details surrounding stock performance during 2017 and 2018 and provide insight into wether or not particular stocks are worth investing in. The "improve" portion of the project was based around improving efficiency of the original code worked on during the module.

**Analysis:** 
Before we were able to pull the analysis for both years we were tasked with refactoring the code given to us. The first step was to import the initial code which gave us input box, chart headers, ticker array, and the ability to activate the worksheet. Then we were able to refactor that code by writting in our additional code (shown below)

'1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
        
        tickerVolumes(i) = 0
        
        tickerStartingPrices(i) = 0
        
        tickerEndingPrices(i) = 0
        
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
         
         For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
             tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
                
             End If
            
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
        
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
             End If
            
            '3d Increase the tickerIndex.
            
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
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

**Results:** In a summary statement address the following:

- The main advantages to refactoring our code would be faster programming and the ability to debug a cleaner code making it easier to fix issues for other people looking at it. The disadvantages to refactoring would not having the initial code set up properly and then refactoring a code that may be faulty, or even having applications that are to large to refactor. 

- Prior to refactoring the code, the execution time to run the stock analysis for 2017 & 2018 was roughtly .8 seconds. Below you can see we were able to decrease that time for both years by roughly .6 seconds


![VBA 2017 Screenshot](https://github.com/paree1kd/Stock-Analysis/blob/9bf72cc59df4268366c4dc7784866d814b66a322/Resources/VBA_Challenge_2017.png)
![VBA 2018 Screenshot](https://github.com/paree1kd/Stock-Analysis/blob/9bf72cc59df4268366c4dc7784866d814b66a322/Resources/VBA_Challenge_2018.png)
