# All Stock Market Analysis

## Oveview 

### Background
The first stock market in the world was the London stock exchange. It was started in a coffeehouse, where traders used 
to meet to exchange shares, in 1773. While today it is possible to purchase almost everything online, there is usually a 
designated market for every commodity. A stock market is a similar market for trading various kinds of assets or company 
shares in a controlled, secure with zero- to low-operational risk and managed environment. While earlier stock markets 
were paper-based, the modern day computer-aided stock markets operate electronically (1). Benifiting from sophisticated 
digital platforms, stock markets prodce huge amounts of data on daily basis compiled from hundreds of thousands of 
transactions. Analysing such data helps understand the total daily volume of shares trded, trends in the market and 
yearly return for each stock. The statistical analysis of such datasets help informed decision making for investment 
in this huge market.

Previously, we prepared a Microsoft Excel Visual Basic macro fro Steve, a finace degree graduate, who would like to help his parentd decide where to invest,especially in the green energy production. They had decided to invest in [Daqo New Energy Corp](http://www.dqsolar.com/company.php) (NYSE: DQ) that manufacturers high-purity polysilicon for the global solar pannel industry (2). In the previous macros we created, Steve was able to to analyze a small number of green energy stocks to compare their performance during 2017 and 2018 but the calculations were slow and were taking noticble amount of time, and even were showing glass hours, that are indicative of a timetaking and labour intensive process by computers. 

### Purpose
 Steve has indicted that he wants to use to VBA macro code we prpare for him to anlyze data from numerous number of stockmarket corporations. Therefore the VB macro needs to be able to analyze large amounts of data in a short amount of time. In doing so, we need to refactor the original script so that it can go through numerous sheets of stoch data and return results faster. To test the VBA  macro we have prepared, we will analyse stock market dataset from trades of different stock market corporations during 2017 and 2018 and compare their perforance. 


## Results
Refactoring the VBA macro, we repurposed the code to be able to analyze large stock market datasets. To test the VBA macro, we performed the nalysis on the green stock dataset that contains data for 12 green energy stocks during the years 2017 and 2018 (Figure 1). The execution times between the original and refactored scrits were also different to a great extent (Figure 2). 

### Stock Market Performance between 2017 and 2018
The stock market performance in 2017 was more noticeable compared to 2018, with only one unsuccessful return outcome in 2017 compared to only two successful return outcomes among 12 green energy stock companies (Figure 1). The largest returns were for DQ and SEDG companies in 2017 but both suffered lossess in 2018, while ENPH and RUN companies kept their returns to a high degree relative to other stocks and RUN even multiplied its return nearly 15 times.
**Figure 1: All Stocks Performance in 2017-2018**
-------
![All stocks 2017-2018 comparison](https://github.com/BHashemi2021/All-Stock-Analysis/blob/main/Resources/All%20stocks%202017-2018%20comparison.png)
------
 
### Execution times of the original script and the refactored script. 
The refactored scripts resulted in a much smaller amount of time compared to the original scripts for both 2017 (0.4 sec vs. 4.1 sec) and 2018 (0.41 sec vs 4.18 sec) (Figure 2). 

**Figure 2: VBA macro execution times betwen original and refactored scripts**
-------
![execution-time-comparisons](https://github.com/BHashemi2021/All-Stock-Analysis/blob/main/Resources/execution-time-comparisons.png)
------
 This noticible difference of up to onetenth of the time between the refactored versus original scripts shows how refactoring and use of r loops and nested loops will improve the efficiency of macro codes (Fig 3).

## Summary

In a summary statement, address the following questions.
1.	What are the advantages or disadvantages of refactoring code?
2.	How do these pros and cons apply to refactoring the original VBA script?

o	There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
o	There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just 
want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to 
make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always 
be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working 
with the existing code at a job (4).



## References

1- Stock Market, Investopedia. Accessed Mat 9, 2021. Retrieved from: https://www.investopedia.com/terms/s/stockmarket.asp
2- DAQO New Enery, Company Introduction. Accessed May 9, 2021. Retreived from: http://www.dqsolar.com/company.php
3- VBA Code To Calculate How Long Your Macro Takes To Run, Chris Newman. January 28, 2015. Accessed May 9, 2021. Retrieved from https://www.thespreadsheetguru.com/the-code-vault/2015/1/28/vba-calculate-macro-run-time
4- What are the advantages and disadvantages of refactoring code smell in software quality?, StackOverflow website. Accessed May 9, 2021. Reterieved from: https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software

