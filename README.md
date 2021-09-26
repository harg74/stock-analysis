# Green Stocks Analysis

## Overview of Project

In order to analyze several tickers for Green Stocks a refactored code was made with the help of Arrays. Where we are able to output from an entire dataset, the Volume transacted and the Return for each ticker, for years 2017 and 2018.

## Results

As we can see in the first image below for 2017, it was a good year for Green Stocks in general. Specially for DQ and SEDG, that has a return of 199.4% and 184.5% respectivaley in one year. In the other hand, TERP had a negative return of -7.2%.

![Returns2017](https://user-images.githubusercontent.com/78564912/134818752-3706bd99-2438-427d-b0f4-1c8943f5cdcf.png)


Regarding year 2018, we can see that it was a bad year for the Green market in general. Where only two companies ENPH and RUN had postive returns of 81.9% and 84.0% respectively. It is important to mention that the most prominent negative returns were for companies DQ with a return of -62.6% and SPWR with a return of -60.5%.

![Returns2018](https://user-images.githubusercontent.com/78564912/134818756-60300525-bf7f-4d55-851f-301268d6d09f.png)


We were able to obtain this kind of results mainly for the below chunk of code. Where we are using arrays to "capture" the values that we need and store them on these arrays.

`   '1b) Create three output arrays
    Dim tickerVolumes() As Long
    Dim tickerStartingPrices() As Single
    Dim tickerEndingPrice() As Single
    
    ReDim tickerVolumes(12)
    ReDim tickerStartingPrices(12)
    ReDim tickerEndingPrice(12)`

For example, to obtain the ticker volumes, we utilized the dynamic tickerVolumes() array and a nested for loop that variable 'i' would loop for each row, looking for the volumes (either in 2017 or 2018 tabs) and 'j' would be looking for the ticker names, where tickerIndex variable would be increasing for each loop 'j' would do.

Inside this nested loop, there are 3 conditionals to determine the values of the tickerStartingPrices() and tickerEndingPrice() arrays. As noted in the images below, the tickerIndex varible is the one indexing inside of these for the four arrays: tickers(), tickerVolumes(), tickerStartingPrices() and tickerEndingPrice().

`  
''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 0 To 11
    
        tickerVolumes(tickerIndex) = 0
    
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            
            Worksheets(yearValue).Activate
            
            If Cells(i, 1).Value = tickers(tickerIndex) Then
                
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8)
                    
            End If
                
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
            
            End If
        
        Next i
        
        tickerIndex = tickerIndex + 1
        
    Next j
    `

Below, are described the scored times for both codes.

   **1. Non-Refactored script:**
    
       a. Analysis for 2017: 0.91 seconds
       ![2017](https://user-images.githubusercontent.com/78564912/134818531-2a267733-5a93-4d48-981f-e44f0abe204a.png)
        
       b. Analysis for 2018: 0.85 seconds
       ![2018](https://user-images.githubusercontent.com/78564912/134818594-0031e939-214f-47bd-af13-99da8d355377.png)


   **2. Refactored script:**
    
       a. Analysis for 2017: 88.32 seconds
       ![2017-refactor](https://user-images.githubusercontent.com/78564912/134818529-32e174d5-d37a-45b7-be63-960d1d6e166e.png)
        
       b. Analysis for 2018: 88.58 seconds
       ![2018-refactor](https://user-images.githubusercontent.com/78564912/134818584-e32ba472-194b-4dd4-831e-58733738971b.png)

## Summary

Refactoring our code should help us to optimize it, once we have finished our first script, to make it easier to read and to make it run faster. Probably it would take you a couple of hours more to refactor it, depending on the circumstances, but at the end it should help you to save minutes, that later convert into hours, while running it in production.

As described above, it took around ≈ 88 seconds more for the refactored code to provide the same results as the previous code. It was certainly more readable, but in terms of speed of execution, at least in this try, it wasn't faster.
    
