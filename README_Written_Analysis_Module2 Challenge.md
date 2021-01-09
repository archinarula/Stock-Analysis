# Written Analysis of the Results for the Stock dataset-Challenge 2 (Refactoring the VBA All Stock Analysis code)

## Overview of Project
We are performing data analysis on Stock ticker data for the year 2017 and 2018 to help Steve, analyse the performance stocks over these 2 years data. We started with analysis of a specific stock (Ticker Symbol "DQ") in order to make recommendations to his parents if it is performning well to invetst. We expanded that single ticker code to also be able to provide Steve with an All Stock analysis code for the 2 years available data, however Steve now wants to use the same code to expand his analysis on more wider stock market data over past few years hence now there is a need to Refactor the original code. 

### Purpose
- Refactor VBA code and measure performance
- Provide a written analysis of the results

## Results
- We refactored the original code to replace the hard coded ticker initialization with a more flexible variable tickerindex 
- We then used this variable to generate output for the arrays tickerVolumes, tickerStartingPrices and ticketEndingPrices 
- We coded the For loop to increase the different output arrays
- We also did the If Then statements to return the ticker performance
- We added the YearValue input box to dynamically allow the users to input year of their choice to generate analysis
- We finally added the Formatting and a Clear Data utility macro buttons to the analysis sheet to display data in a more organized and visually appealling manner

![VBA_Challenge_Year Input] (https://github.com/archinarula/Stock-Analysis/blob/main/VBA_Challenge_Year%20Input.jpg)
![VBA_Challenge_2017] (https://github.com/archinarula/Stock-Analysis/blob/main/VBA_Challenge_2017.jpg)
![VBA_Challenge_2018] (https://github.com/archinarula/Stock-Analysis/blob/main/VBA_Challenge_2018.jpg)


## Summary

### Advantages of refactoring
- The code is robust and flexible to accomodate a wider dataset without having the need to redo the code for a similar analysis requirement
- The code is more organized and clear on sections
- It is easier to debug
- It is easier to reuse and customize section wise

### Disadvantages of refactoring
- Was time taking as it was required to add quite a bit of more logic to complete the code
- Risk of mistakes while reading and editing the correct pieces of code for accuracy on overall script so needs a detailed view and skills
- Debugging to make all lines of code work with code breaks was required to achieve the results

### How these Pros and Cons apply to refactoring the original VBA script
- It was helpful to reuse pieces of code done previously rather than starting from scratch
- Biggest advantage of refactoring on this case was that the original code was not complete so while it was working it was good for a restricted number of data (tickers), after refactoring the analysis could be completed and expanded to a wider dataset over multiple years
- Performance time of refactored code was better than original
- There was no disadvantage to refactoring on this case



