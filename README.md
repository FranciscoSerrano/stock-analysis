# VBA of Wall Street

## Overview
### Purpose
The purpose of this analysis is to help Steve in choose stocks for his parents to buy based on data from 2017 and 2018.

### Background
I had previously assisted Steve by creating a Workbook with some VBA macros that would parse through the data and produce a chart with data he could draw meaning from. After reviewing the macro I then wondered if the VBA could be made more efficent. The first program took **0.578125** seconds to run for the 2017 data set and **0.546875** seconds for the 2018 data set. My goal is to refactor the code so that it may run more efficiently. The results for this data set might not be significant but when steve wants to look at the entrire market I will make a substantial difference.

## Results

### Analysis
Steve wants to find out what stocks his parents should invest in based on the data from 12 companies 2017 and 2018 returns.  

In order to help Steve, I created a sheet using VBA that includes a table which shows the return of each stock per year. The table uses conditional formatting to show which stocks have a positive rate of return and which stocks have a negative rate of return. Positive returns have a green cell and negative returns are colored red.   

Based on the results, Steve should recommend that his parents invest in the ENPH and RUN stocks because they are the only two stocks listed that had a positive return in both 2017 and 2018.

![2017 Stock Data](/resources/2017_results.png)

![2018 Stock Data](/resources/2018_results.png)

### Refactoring
Steve mentioned he wanted to look at the whole stock market for these two years. Aftering hearing this I wondered if I could make my VBA a little more effeciant realizing it'll be parsing a much larger data set.   

During refactoring I set the index to 0 before going further. I also made a stand alone loop to set the volume of each stock to zero so that the program would not have to do it on each iteration. It was also changed so that the loop stayed on one sheet before moving to the next loop and did not have to switch back and forth each time.

```VBA
   
    tickerIndex = 0

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i

    Worksheets(yearValue).Activate
    For i = 2 To RowCount
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If

            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                tickerIndex = tickerIndex + 1
                
            End If
        
    Next i

    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

```

After these changes where made there was an time efficency gain that was an **order of magnitude** faster than the orignal time.

![2017 VBA Data](/resources/VBA_Challenge_2017.png)

![2018 VBA Data](/resources/VBA_Challenge_2018.png)


## Summary
### What are the advantages or disadvantages or refactoring code?
The advantge to refactoring is that you can find areas in your code where tasks are repeated or the same variables are used more than once so you can come up with ways to consolidate these findings to make the code run more effectively. This can result in faster runntimes and less crashes, bugs, or errors.

A disadvantage of refactoring code is that it takes time. There is a cost to benefit when it comes to investing that time. Is it worth an hour, day, or week of work to save seconds? Sometimes it might be other times it might not.

### How do these pros and cons apply to refactoring the original VBA script?
The pros applied when removing the loops made the script much quicker as shown in the photos used for the analysis. It also made the script much easier to understand if a colleague were to go in and refector the script further.

The cons occurred when I realized how much time I spent to save a thousandth of a second on a program that wasnt that slow to begin with. Although I can undertsand how it might be useful with much large data sets. A 10x decrease in time is the difference beween a day and just a couple hours.
