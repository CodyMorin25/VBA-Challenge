# VBA-Challenge

## Overview of Project
The location of the excel sheet is: [github repository](https://github.com/CodyMorin25/VBA-Challenge.git)

### Purpose and Background
The purose of this project was to refractor Excel VBA code that collected info on stocks in years 2017 and 2018, in order to determine if they were worth investing in. The original format was done in a similar matter but was refractored inorder to make it easier to read and more effecient.

## Results
### Analysis
In my analysis before refactoring I copied code needed for this project including code to create input box, headers, ticker arrays, and to make sure the appropriate worksheet is being activated. Then the necessary steps were listed to give structure to the refactoring. Below is the code and instructions.


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
            
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
    
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            
            End If

            '3d Increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            
                tickerIndex = tickerIndex + 1
                
            End If
            
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("AllStocksAnalysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i


## Summary
### Advantages and Disadvantages to Refactoring
Refactoring helps make the code easier to read (i.e. less complex and easier to maintain). However refactoring can be time consuming, there is no way to be sure how long it will take to complete the process. Also you may need to retest a lot of functionality.

### Advantages and Disadvantages to Original or Refractored VBA script
The biggest benefits to come from refactoring is the macro run times and how much easier it is to read than the original. After comparing the two I found the refractored code much easier to follow. In the case of the macro run times originally it was taking close to one second to run but now it is running in a hundredth of a second. Below are screenshots of the refractored run times for each year.



