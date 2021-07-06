# GW Module 2 Excel and VBA
## Overview of Project

Take the challenge starter code that was developed for Steve and refactor it to run more efficiently. Produce analysis multiple stocks simultaneously in the time frame of 2017 and 2018 to highlight which stocks are performing positively and negatively over time. 

## Results

### Stock Performance
Stocks in 2017 had greater returns with the exception of "TERP". The only stocks that performed positively in 2018 were "ENPH" and "RUN" which also were more than 200% increased in terms of Total Daily Volume metrics. This indicates that these companies were heavily perceived by the trading population to be hot stocks. 

### Code Performance
Refactored code was able to crunch data in a little over half a second compared to the original starter code provided in this challenge. I will mention that the processing time improvements were minimal in this exercise, from what I could see. This probably has to do with revisions not addressing other inefficiencies in how the loop executed iterations across the 3000+ rows of data.

### All Stocks Analysis Tables and Process Time

![2017 Analysis](https://github.com/demarcomf/stock-analysis/blob/main/VBA_Challenge_2017.PNG)
![2018 Analysis](https://github.com/demarcomf/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

## General Summary
- General advantages of refactoring code: on a simple level, should improve code run times and traceability if this were a corporate team or organization. 
- General disadvantages of refactoring code: breaking code or not saving the original form. Utilizing detailed comments make for an easier time following the logic chains when refactoring and improving code. 

## VBA Refactoring Specific Summary
- Advantages: This provided a fundamental layout of the data processing needed as we aimed to search for particular stocks and the values associated with them over 2017/2018.
- Disadvantages: Even with adding all the arrays (tickerVolumes, tickerStartingPrices, tickerEndingPrices), I did not notice a significant run time improvement. It was a lot of effort for saving half a second. I probably could have refactored this better.
