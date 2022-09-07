# Stock Analysis

## Overview of Project
In this project we are working to help our friend Steve analyze some stock data in the green energy sector. Steve's parents are very interested in investing and have come to him to invest all their money into DQ, a company that makes silicon wafers for solar panels. Steve however, things that they should diversify more and have asked us to provide some actionable analysis to prove if his theories are correct.

### Purpose
The purpose of this analysis was to provide usable data to Steve in order to determine if DQ is a good investment or if there should be more diversity in his parent's portfolio.

## Analysis and Challenges
In the data set provided we look at the Total Daily Volume and Annual Return for the years 2017 and 2018 for 12 different stocks.

The first stock we analyzed was the DQ trends for 2017 and 2018.

We created some VBA macros that allowed us to populate the Year, Total Volume and Return for DQ in 2017 and 2018. The results were surprising.

![DQ_Analysis](/Resources/DQ_Analysis.png)


As you can see from this image, DQ did not perform well at all. Steve decided at this point that he wanted to be able to quickly compare all 12 of the stocks to determine the best performing stocks to suggest for his parent's portfolio.

In order to make the process more accessible we created a new macro assigned to a button that allowed Steve to input the year needed and returned the same data analysis for all stock tickers in the spreadsheet.

        yearValue = InputBox("What year would you like to run the analysis on?")

Then by assigning the tickers to an array we looped through all the tickers in the spreadsheet to determine the Total Volumes and Returns.

        'Assign the tickers to elements of the array
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

We encountered another challenge during this process. The runtime for the data analysis seemed to be rather slow. Steve requested a timer for us to determine the overall run time of each set of data.

        Dim startTime As Single
        Dim endTime  As Single

            startTime = Timer
            

            endTime = Timer
    
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

This code allowed us to display the runtime of the data analysis for each year at the end of the process.

![Pre_Refactoring_2017](/Resources/Pre_Refactoring_2017.png)
![Pre_Refactoring_2017](/Resources/Pre_Refactoring_2018.png)

As we can see from the above pictures the code did not take that long to run however the concern was as the data grew with more stocks added it could become quite cumbersome to research. We decided to refactor the code and use more arrays to expedite the analysis process. After refactoring the code by creating arrays for the outputs of tickerVolumes, tickerStartingPrices, and tickerEndingPrices as well as several other improvements to the code our runtimes were improved significantly.

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

### Analysis of Outcomes 
In the outcome tables below, we determined that the two stocks with the highest returns for 2017 and 2018 were ENPH and RUN.

![Stock_Data_2017](/Resources/Stock_Data_2017.png)
![Stock_Data_2018](/Resources/Stock_Data_2018.png)

You can see from the data that while DQ had a positive return in 2017 they fell off significantly in 2018.

### Summary

As shown in the previous examples of the pre refactored code vs refactored, there is a significant increase in the runtime of the process after refactoring. While the refactoring of code did require more research to determine the best practices, the potential outcome of increased productivity during research proved to be enough to justify the code enhancement.

The original VBA script vs the refactored code was also much bulkier than the final product of the refactored code.

