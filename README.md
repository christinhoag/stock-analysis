# Stock Analysis

## Overview of Project

### The purpose of this project was to analyze whether certain stocks in 2017 and 2018 were profitable to invest in. We refactored existing code in VBA Excel to hopefully run faster and more efficiently if Steve should choose to analyze thousands of stocks in the market.

## Results

### In 2017, eleven out of twelve stocks evaluated had a positive return. The two top performing stocks were DQ with a return of 199.4%, and SEDG with a return of 184.5%. TERP was the only stock that had a negative return (-7.2%). In 2018, it seems that the market took a dip as only two out of twelve evaluated stocks yielded a positive return. Those two were ENPH with a return of 81.9%, and RUN with a return of 84%. The worst performing stock was DQ with a return of -62.6%.

### In the refactored code we created a ticker index to loop through rather than simply hard coding which ticker to pull stats from. This runs the macro to find the starting and ending price of each stock all at once rather than going through all the tickers one at a time. The refactored portion is shown below.

![refactored_code](https://user-images.githubusercontent.com/104405357/168727415-8e10a41e-bac3-451a-959a-39160508d4c9.png)

### Our refactored code slightly decreased run time on the macro by 2/10ths of a second for both years 2017 and 2018. Attached below are the screenshots demonstrating the new run times.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/104405357/168727629-f5a1a272-8a96-4aa5-b402-3b3c115a1a7c.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/104405357/168727645-68294075-24cc-4a30-a281-eb3001515ee5.png)

## Summary

### Refactoring code can be advantageous in some cases because it means not having to write the code from scratch, but rather editing the script to perform exactly how you want. In the refactoring process you can change the code to be more efficient, dynamic, and organized so others can use and read it easier. However, the biggest disadvantage is that it can be very time consuming to debug any mistakes made along the way. I experienced this disadvantage firsthand while refactoring the All Stocks Analysis code. It took me approximately 4 hours of debugging to finally get the macro to run. However, once it ran properly it did, in fact, run quicker than the original script from the Module 2 lesson; proving that refactoring can result in more efficient code.
