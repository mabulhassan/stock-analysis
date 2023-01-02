# stock-analysis
## Overview of Project
### Purpose
The purpose of this project is to analyze data for stocks of two consecutive years and determine if the investment value. This is accomplished by refactoring VBA code to make it run faster
and more effecient in order to generate the information needed for analysis.


## Results
### Analysis
After running the analysis, it was apparent that year 2017 had better performance than following year 2018. This was displaying using the red colors and reading the positive performance percentage
of the previous year.
Refactoring the VBA code by removing the formatting section and including it in the same iteration that outputs the results made the program run faster. This is evident from the time taken
to run the program before and after the refactoring.
### Before Refactoring
2017 -   1.13 seconds
2018 -   0.15 seconds
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
### After Refactoring
2017 -  0.1 seconds
2018 -  0.093 seconds
 'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
       
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        If Cells(4 + i, 3).Value > 0 Then
              Cells(4 + i, 3).Interior.Color = vbGreen
         Else
               Cells(4 + i, 3).Interior.Color = vbRed
         End If
        
    Next i
  

## Summary
### Pros and Cons of Refactoring Code
Refactoring in general makes code more organized and easier to read and run more effeciently. Disadvantage is the time it takes to modify and test the code and risk of breaking a functioning code.

### The Advantages of Refactoring Stock Analysis
In the case of the VBA code, the code ran faster after refactoring. It took more than 1 second to run the code for 2017 and 0.15 second for 2018 prior to refactoring.
It's evident from the images captures prior and after the code refactoring that the time to run the program decreased to 0.1 seond and 0.093 seconds. However, the disadvantages was the time it took to restest.
### Before Refactoring
![VBA 2017 Screenshot](https://github.com/mabulhassan/kickstarter-analysis/Resources/VBA_Challenge_2017.PNG)
![VBA 2018 Screenshot](https://github.com/mabulhassan/kickstarter-analysis/Resources/VBA_Challenge_2018.PNG)
### After Refactoring
![VBA 2017 Screenshot](https://github.com/mabulhassan/kickstarter-analysis/Resources/VBA_Challenge_2017New.PNG)
![VBA 2018 Screenshot](https://github.com/mabulhassan/kickstarter-analysis/Resources/VBA_Challenge_2018New.PNG)
