# Green Stocks Analysis

## Overview of Project

In order to analyze several tickers for Green Stocks a refactored code was made with the help of Arrays. Where we are able to output from an entire dataset, the Volume transacted and the Return for each ticker, for years 2017 and 2018.

## Results

As we can see in the first image below for 2017, it was a good year for Green Stocks in general. Specially for DQ and SEDG, that has a return of 199.4% and 184.5% respectivaley in one year. In the other hand, TERP had a negative return of -7.2%.

![Returns2017](https://user-images.githubusercontent.com/78564912/134818752-3706bd99-2438-427d-b0f4-1c8943f5cdcf.png)


Regarding year 2018, we can see that it was a bad year for the Green market in general. Where only two companies ENPH and RUN had postive returns of 81.9% and 84.0% respectively. It is important to mention that the most prominent negative returns were for companies DQ with a return of -62.6% and SPWR with a return of -60.5%.

![Returns2018](https://user-images.githubusercontent.com/78564912/134818756-60300525-bf7f-4d55-851f-301268d6d09f.png)


We were able to obtain this kind of results mainly for the below chunk of code. Where we are using arrays to "capture" the values that we need and store them on these arrays.

    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrice(12) As Single

For example, to obtain the ticker volumes, we utilized 1) the dynamic tickerVolumes() array, 2) a for loop with variable 'i' that would loop for each row, looking for the volumes (either in 2017 or 2018 tabs) and 3) another for loop with variable 'j', that  would be setting the tickerVolumes() to zero.

Inside the first loop there are 2 conditionals to determine the values of the tickerStartingPrices() and tickerEndingPrice() arrays. As noted in the images below, the tickerIndex varible is the one indexing inside of these for the four arrays: tickers(), tickerVolumes(), tickerStartingPrices() and tickerEndingPrice().

```

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 0 To 11

        tickerVolumes(j) = 0

    Next j
    
    '2b) Loop over all the rows in the spreadsheet.
     Worksheets(yearValue).Activate
     For i = 2 To RowCount
        
            '3a) Increase volume for current ticker
            
           tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
                
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
            
            '3c) check if the current row is the last row with the selected ticker
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
                
                tickerIndex = tickerIndex + 1
            
            End If
        
        Next i
```


Below, are listed the scored times for both codes:

### 1. Non-Refactored script:
    
   - Analysis for 2017: 0.91 seconds
    
        ![2017](https://user-images.githubusercontent.com/78564912/134818531-2a267733-5a93-4d48-981f-e44f0abe204a.png)
        
   - Analysis for 2018: 0.85 seconds
       
       ![2018](https://user-images.githubusercontent.com/78564912/134818594-0031e939-214f-47bd-af13-99da8d355377.png)


### 2. Refactored script:
    
   - Analysis for 2017: 0.17 seconds
    
        ![2017 Refactored](https://user-images.githubusercontent.com/78564912/136880060-eee8957e-17ce-4a5c-833e-25307525432e.png)

        
   - Analysis for 2018: 0.17 seconds
    
        ![2018 Refactored](https://user-images.githubusercontent.com/78564912/136880074-65f3cbe8-87a8-41a8-b152-f148a75279ec.png)


## Summary

Refactoring our code should help us to optimize it once we have finished our first script, to make it easier to read and to make it run faster. Probably it would take you a couple of hours more to refactor it depending on the circumstances, but at the end it should help you to save minutes, that later convert into hours, while running it in production.

As described above, it took only around ≈ .17 seconds for the refactored code to provide the same results as the non-refactored code. Its benefits, it certainly more readable, and in terms of speed of execution, it is faster.
    
