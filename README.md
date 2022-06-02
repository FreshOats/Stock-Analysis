# Stock Analysis Using VBA
***A comparison of a nested loop and refactored code to eliminate the nesting to improve efficiency***
#### by Justin R. Papreck
---

## Overview
  In this analysis using VBA in Excel, a client is trying to determine which of 12 to invest in based on data from 2017 and 2018. This project employs a subroutine written in VBA to take an input year, calculate the total daily volume and the return from starting price to ending price for each of the tickers.  
  
### Purpose
  The purpose of the project is to reduce the processing speed by refactoring the code that used a nested for loop. While the processing time for the nested loops weren't terrible, they were only analyzing 12 different tickers. Increasing the number of tickers to hundreds would seriously impact the speed and resources required to count everything within a nested loop. The code was therefore refactored to use a single loop with conditional statements within to determine when to change to a different ticker. 

---
## Results and Analysis
  For each of the years 2017 and 2018 the refactored code ran between 7 and 8 times faster than the original code using nested loops. 

---
### Analysis of the Nested Loop code
  The main difference between the original code and the refactored code was the use of a nested loop, going through all of the rows for each ticker. While this works, it also means that the program is attemtping to analyze all of the rows of the dataset, even after the rows for that ticker have ended. In essence, it is analyzing 11 sets of additional rows 11 additional times, so about 121 times more than the actual required analysis, assuming there are equal numbers of ticker entries. The time of execution for the 2017 and 2018 datasets are shown below: 

![Nested_2017](https://user-images.githubusercontent.com/33167541/171732305-8b0e1b55-ccb4-42c4-b444-fcdc824299e8.png)
![Nested_2018](https://user-images.githubusercontent.com/33167541/171732321-f9cd7723-3792-4540-866e-7eaf522597a1.png)

While 0.625 seconds doesn't seem like a long time, this code is only evaluating 12 different stocks, when in reality, a client would be expecting the analysis of hundreds. Refactoring may not be necessary, but if we were to scale up the analysis, it could make a significant impact on the processing time as well as the computing resources necessary to do the analysis.   

The nested loop used in this analysis is as follows.
```
'Loop through the tickers

 For i = 0 To 11
        
    ticker = tickers(i)
    totalVolume = 0

    'Loop through the rows in the data
     Worksheets(yearValue).Activate
        For j = 2 To RowCount
    
         'Find total volume for current ticker
            If Cells(j, 1).Value = ticker Then
                
                totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
                
         'Find starting price for ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            
                startingPrice = Cells(j, 6).Value
            
            End If
    
         'Find end price for ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                
                endingPrice = Cells(j, 6).Value
                
            End If
         
        Next j
        
'Output data for current ticker
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = ((endingPrice / startingPrice) - 1)
                 
Next i
```

The ouputs for each of the years were: 
![Nested_Loops_2017](https://user-images.githubusercontent.com/33167541/171733814-ae6db10e-d5d9-45f6-ab5c-0687a8e61400.png)
![Nested_Loops_2018](https://user-images.githubusercontent.com/33167541/171733846-9b50539d-5b2f-4083-b214-462a597ada22.png)

In refactoring this code, it is possible to eliminate the need to loop through ALL of the rows of the dataset for each ticker by using the identification of the ending of each individual ticker. By adding an additional IF statement when the condition to determing the endingPrice is called, the ticker can be changed using indices, which would allow the program to continue looping though the rows, but only needing to do it once. 

---
### Analysis of the Refactored Code
  The initial change in the code took place in the creation of arrays for the volumes, and starting/ending prices. The volumes needed to be looped through and initialized to zero before starting the analysis because this was part of the outer loop in the nested loop function. The next part of the loop no longer needed the nesting. As the program runs through each row, it encounters an If statement to determine wither it is a starting row, where the starting price is recorded, or an ending row, where the ending price is recorded. However, since it recognizes the ending row as the last row of that ticker, a statement is added to increase the ticker index by 1. Now as the next set of rows are looped through (respectively analyzing the next ticker from its start to finish), the totals are only added to the volumes for that ticker. 

```
'Initializing the tickerVolumes to zero.
        'Creating this loop to set up the tickers and their volumes outside of the other loop
        'The previous example nested these loops, which was slower
    
    
    For j = 0 To 11
        
        ticker = tickers(j)
        tickerVolumes(j) = 0
        
    Next j
```

```
'' Loop over all the rows in the spreadsheet
    For i = 2 To RowCount
        
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            
        'Check if the current row is the first row with the selected tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
                tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
                'This just records the starting value for the Return calculation
        
            End If

        
        'check if the current row is the last row with the selected ticker
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
                'This just records the ending value for the Return calculation
            
          
        'If and only if the script determines the ticker to be at the end of its row, this will add 1 to the ticker value
         'and thus changing to the next ticker, this is what allows us to process through the rows a single time
        
                If tickerIndex < 11 Then
                    
                    tickerIndex = tickerIndex + 1
                    
                End If
                
             End If
        
    Next i
    
```
  
Ultimately, the program only analyzes all of the rows from the dataset once, rather than 12 times. The resulting runtimes are as follows: 
  
![VBA_2017](https://user-images.githubusercontent.com/33167541/171735415-c5f46b4e-5fb5-4b95-b428-989d74ea830c.png)
![VBA_2018](https://user-images.githubusercontent.com/33167541/171735437-3098c28c-64cf-472f-8891-9f97076220e1.png)

As can be seen, these times are on a magnitude of a tenth of the times with the nested loop. The results of the analysis are the same as the original analysis:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/33167541/171736228-cff56885-1aa0-45bd-a660-67dc703d8c17.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/33167541/171736249-c9f55485-fa37-4b5f-9d35-21895aa42a1a.png)

---
### Analysis on the Data
  As the initial intent of the project was the analysis of stocks, given the different tickers to follow for the years 2017 and 2018, there were 3 significant tickers in the findings. Both ENPH and RUN had consistent gains in both 2017 and 2018. TERP, on the other hand, was the only stock to have losses in both years. If the client were to only pick one stock to invest in, the ENPH had a gain of 130% in 2017 and an 82% gain in 2018, whereas RUN only had a 6% gain in 2017 despite a similar 84% gain in 2018. The client was initially investigating DQ, which had an impressive 200% gain in 2017, but then a 63% loss in 2018. Given these data, the best option at that time was ENPH. 

---
## Discussion
  There are definitely advantages to refactoring code - primarily the refactored code makes the program more efficient, requiring less time and often requiring less processing power. The biggest disadvantage of refactoring code is that it often makes the code more difficult to understand - it's not necessarily as obvious what the program is doing. While refactoring code, a programmer is less interested in readability and more interested in functionality. This also creates more of an issue with debugging the program when it fails to produce an expected outcome. As we reduce the complexity of the code to improve efficiency, we lose the simplicity in seeing how each step directly impacts the outcomes. 
  
  In VBA, the refactored script required the changing of some variables to arrays, something that wasn't inherently obvious. Unlike other programming languages, it is not easy to determine the data types of each of the saved variables, so tracing the data as it is processed isn't as visible, so when trying to debug issues as they arise, addressing these roadblocks was a challenge. Fixing the code with the nested loop was much easier than trying to debug the more efficient refactored code, mostly due to needing to fully understand the data types and data flow, which is still unclear. The advantage to refactoring the script was obvious - it did make the processing speed almost 8 times faster. For the purposed of this analysis, it wasn't necessary.
  
  Therefore, one takeaway is that the code is probably best in its most simplistic format until it needs to be refactored. In the analysis of 12 stock tickers, our runtime with the nested loops was still under 1 second. The refactored program cut that down to less than 1/10th of a second, but for the purposes of our analysis had no impact. Had we scaled up the dataset to incorporate hundreds of tickers, the impact may have deemed the refactoring necessary. 
