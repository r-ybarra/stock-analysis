# stock-analysis

##Overview of Project: Explain the purpose of this analysis.

Analyizing stock trends performance to discover investment opportunities.

##Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Comparison
![2017_results](https://user-images.githubusercontent.com/83841580/123562571-1ba6e500-d775-11eb-9088-66d7226150c3.png)
![2018_results](https://user-images.githubusercontent.com/83841580/123562572-1ea1d580-d775-11eb-8123-1cd41766e2b9.png)

Comparing the results you can see that ENPH Made a profit in both years, but a lesser profit in 2018. RUN also made a profit in both years, but increased in profitablility in 2018. All other companies lost money in one or both years.

### Original Code
```
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i
```
![old_2017](https://user-images.githubusercontent.com/83841580/123562509-bfdc5c00-d774-11eb-9fe4-6f2d6db9ec14.png)

This version uses nested for loops and an output during each iteration of the first loop. The inner loop checks the entirety of the data during each loop which wastes time and resources.

### Refactored Code
```
    For i = 2 To RowCount
 
        '3a) Increase volume for current ticker
         tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1) = tickers(i)
        Cells(i + 4, 2) = tickerVolumes(i)
        Cells(i + 4, 3) = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
```
![VBA_Challenge_2017](https://user-images.githubusercontent.com/83841580/123562549-f4501800-d774-11eb-8736-607d0c980ea1.png)

This version only reads the data a single time and instead switches through each ticker sign when necessary. It then uses a small loop to output from memory. This significantly reduces the amount of data that must be read to do the same calculations. You can see from the screenshot that the calculation time was greatly reduced.

## Summary: In a summary statement, address the following questions.
In order to diversify your stock portfolio you should invest in RUN and ENPH. Both of these companies were profitable in 2017 and 2018. RUN is the safest bet as it is trending upward over this time frame.

### What are the advantages or disadvantages of refactoring code?
Any time that you change your code whether that be refactoring or just adding functionality, you expose yourself to the risk of breaking something that was working. On the other hand, when you refactor you have a chance to make it run better and reduce the time it takes to run that code.

### How do these pros and cons apply to refactoring the original VBA script?
In this case I was able to refactor the code to greatly reduce the amount of time it takes to execute this code. I certainly created bugs that had to be squashed along the way, but it worked out in the end.

