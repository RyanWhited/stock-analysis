# Stock Analysis with VBA
## Overview of Project
### Purpose
The purpose of this project was to analyze two sets of stock data to determine their volumes, return rates, and highlight those with positive return rates. I was given the data for 12 different stocks from the years of 2017 and 2018. I initially created a successful project that executed the analysis correctly but was challenged to speed up the results by refactoring (editing) my VBA code. 
## Results
After refactoring the original code, I noticed a significant change in the speed of the output as seen below:
- 2017

![VBA_Challenge_2017_Original](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_2017_Original.png)   ![VBA_Challenge_2017](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_2017.png) 

- 2018

![VBA_Challenge_2018_Original](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_2018_Original.png)   ![VBA_Challenge_2018](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_2018.png)

What helped the process speed up? One of the main reasons is I utilized arrays for the ticker volumes plus both starting and ending prices (see image below). In the original code, we used a nested loop that would obtain these values which slowed down the speed of the results. 

![VBA_Challenge_New_Arrays](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_New_Arrays.png)

In addition, in the original code the output was built into the main loop but in the refactored code I created a separate loop that would loop through the arrays to output the Ticker, Total Daily Volume, and Return.

![VBA_Challenge_New_Output](https://github.com/RyanWhited/stock-analysis/blob/main/VBA_Challenge_New_Output.png)


## Summary
Refactoring the code makes it easier to read, takes up less memory, and executes more effeciently. However, as the saying goes "If it ain't broke don't fix it!" Refactoring code could potentially lead to execution issues if you are not careful. 

In my case as it relates to this project, I definitely ran into some issues (mainly around executing the arrays appropriately) that had to be resolved while refactoring the code. After some trial and error, I finally ended up with a successful macro that sped up the output process significantly. 
