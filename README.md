# Module 2 Challenge

## Overview of Project

### Background
In the Module 2, we got introduced to Visual Basic for Applications so we can get more analytic and complex calculations by automating simple tasks. 
Our motivation is help Steve's parents and him to analyse Green Stocks. In the application, we studied Daqo's stock, DQ, and then we used the structure created to create a flexible macro for running multiple stocks in different years.

### Purpose
In this Challenge, is intended to expand the dataset to include the entire stock market over the last few years. To do that, we want the most effectiveness code so it can run with thousands of stocks.
To accomplish that, we are going to **refract the current code**. In other words, the goal is restructuring the existing *AllStocksAnalysis Subroutine* so we can loop through all the data **one time** in order to collect the same information as we did in the though the Module Orientation.


## Results

### Stock performance between 2017 and 2018

We can clearly see that the **Green Stocks** we analyzed **had a better year in 2017**; in 2018, almost every stock had negative returns.

| 2017 | 2018 |
|--|--|
| ![All Stocks 2017](https://user-images.githubusercontent.com/90414330/135783276-2b31526d-cc7f-449d-9acd-10c3986d4c78.png) | ![All Stocks 2018](https://user-images.githubusercontent.com/90414330/135783291-633043fa-9304-4151-ba92-c9f39f75e50d.png)


#### The Most Brilliant Stock
***ENHP* was the one that had a brilliant performance both years.** *ENHP* not only had positive returns both years: from an amazing growth of 129.5% in 2017, they also managed to grow an additional 81.9% in 2018.

#### A Possible Relationship
Alternatively, *RUN* also kept a positive return both years. Which makes me wonder what do they have in common?

I could notice that the ***Total of the Daily Volume*** had a significant increment one year from another in both stocks: *ENHP* had 221M in 2017 and it increased to 607M in 2018 (almost three times); whereas *RUN* had 267M in 2017 to 502M in 2018 (the double of trades). On the contrary, the rest of the stocks kept similar amounts of ***Total of the Daily Volume*** each year.

We must remark that DQ doesn't align with that observation; those stocks were actively traded three times more in 2018 than 2017, and despite that they had a big decrease of 62.6%. This may be interesting to analyze it further.

#### Don't believe in anything until you can prove it
In adition, we kept in mind that Steve's parents *"believe that if a stock is traded often, then the price will accurately reflect the value of the stock"*. This isn't always true: The stock most traded in 2017 was *SPWR*,  if we had invested in that stock at the end of the year or beginnings of 2018; we could have lost a 44.6% by the end of 2018.
<br> <br> <br>

### Execution times
| Original Code| Refactoring the Code |
|--|--|
| ![Module_2017](https://user-images.githubusercontent.com/90414330/135787208-59d76cf3-6264-4bf4-90e2-e735ec3a4be1.png)<br>![Module_2018](https://user-images.githubusercontent.com/90414330/135787207-fccbf5ef-7871-4538-9b28-ba6ff10b968b.png)| ![VBA_Challenge_2017](https://user-images.githubusercontent.com/90414330/135786417-b1b35f2f-dc9d-4ba6-8a6b-c8dad5b6acfb.png) <br> ![VBA_Challenge_2018](https://user-images.githubusercontent.com/90414330/135786422-f91b6809-6176-4281-8b8c-a58a1a7c1698.png)

It's quite obvious that the refracted code runs faster that the original one, we'll get deeper in the summary.  

Additional, we can notice that with the data of 2017, we get the analysis before; maybe as a result that in 2017 there are less trades per day.


## Summary
1.  **What are the advantages or disadvantages of refactoring code?**

**Pros**
We can always improve a code if we give it a second or even a third look. We can simplify processes, decrease the use of memory, improve the logic, clean the code so it can be more understandable and flexible, and more.

**Cons**
In contrast, if we are not careful, we can break the code and forget about some outputs and elements that might be important for the integrity of the analysis and the code, or end with a dirty code.


3.  **How do these pros and cons apply to refactoring the original VBA script?**

**Pros**
We simplified the storage of the outputs. Instead of going back and forward to the `Analysis Worksheeet` and the `yearValue Sheet` printing the outputs each time we get them, the outputs were storage in arrays, keeping the `yearValue Sheet` active, and then printed using a simpler loop. Which can also be helpful if we want to use the `tickerVolumes`, `tickerStartingPrices`, and `tickerEndingPrices` further on.

**Cons**
In this case, we ended with a somewhat dirty code, but only because of minor details: we start variables in the middle of the code and added an additional loop (to initialize the ` tickerVolumes ` ),
