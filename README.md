# Stock Analysis with VBA

## Overview

The purpose of this project was to use VBA Macros to create code for analyzing stock data in Excel. The idea was to find the return rate of any stock in the table. This macro was originally created for a client to find data on stocks in a specific table, and was refactored so that it would not be limited to just one table, but any stock found on multiple worksheets. I looked specifically at 12 different tickers from our table of data, and found the return rate from 2017 and 2018.

## Results

With my refactored code, I found the rate of return for 12 different stocks and was able to compare data on those 12 stocks between 2017 and 2018. Overall 2017 was a much better year for these particular stocks. This could be an overall pattern in stock data for the year, but hard to tell until a larger set of data has been looked at.

### 2017

<img width="1511" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/105175961/198857478-1b61aea8-3a6a-4fc8-bc18-6f948fabf429.png">

### 2018

<img width="1504" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/105175961/198857482-ac260f91-1207-48be-b44f-31cb08e8f909.png">

### Example of code with variables created for tickers

<img width="915" alt="Screen Shot 2022-10-29 at 6 35 56 PM" src="https://user-images.githubusercontent.com/105175961/198857597-5866eb67-7da6-45a8-93bd-5f02fb763652.png">

## Summary

# Advantages and Disadvantages of Refactoring

Refactoring the code was more advantage than disadvantage as the refactored code is easier to use on a wider variety of data, in this case multiple tables and tickers. It does run faster as well but we are talking about the difference between 0.25 seconds (original script) and 0.05 seconds (refactored script) for this particular instance. The takeaway here is that the refactored code runs slightly more efficiently, and part of that is due to the way it is written. The only disadvantages I can really see is that creating variables for all the tickers used can be a bit time consuming, and more arrays and variables have the potential to create more errors.





