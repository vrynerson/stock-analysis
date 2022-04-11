# VBA Module 2 Challenge
## Overview
The client wanted to be confident in their financial consultation for investing money in alternative energy stocks based on the years 2017 and 2018. The client wanted to diversify their clients' portfolio instead of letting them invest all their money into DAQO. The purpose of this project was to refactor code to accommodate thousands of stocks in a timely manner. With the click of a button, the client is able to see the amount of seconds it took to calculate the total daily volume and return of the input year.  

## Results
### Stock Analysis
Between 2017 and 2018, 2017 was a better year for alternative energy. All of the companies that were a part of this analysis, besides the company with the ticker TERP, had a return in some capacity. The companies with the lowest and highest returns were RUN with a 5.55% return and DQ with a 199.45% return, respectively. If another year was not analyzed, the client would probably have chosen to invest all their money in DQ. Looking at 2018 tells the client that would not be a wise decision. The return on DQ was in the red with -62.60% return, which is the lowest return of all the alternative energy companies. 2018 was not a year for alternative energy. All but two companies' returns were in the negative. RUN and ENPH had returns of 83.95% and 81.92%. Overall, RUN and ENPH would be the best companies to invest in based on 2017 and 2018 dataset. 

Below are tables displaying the total daily volume and return for each ticker for 2017 and 2018. 

<p align="center">
  <img src="https://github.com/vrynerson/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true" width="300" height="455">
  <img src="https://github.com/vrynerson/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true" width="300" height="455">
 </p>


### Refactor Observations
The amount of time it took to process the data is also shown in the images above. With the refactored code, the 2017 data was processed in 0.23 seconds and the 2018 data was processed in 0.17 seconds. Before refactoring the code, the code for 2017 ran in 1.53 seconds and the code for 2018 ran in 1.62 seconds.

## Summary
### Refactoring Code Pros and Cons
The advantage of refactoring code leads to a more conscise process, which in turn, leads to a faster process. When working with multiple people on the same coding project, it could be a more positive and efficient experience when collaborating. A disadvantage could be taking the extra time to reconstruct a code that worked well in the first place. It may be that you are unable to refactor the code as you wish it would be and that time could have been spent on a different project. 

### Refactoring Details
The refactored code yields and average of an 87% decrease in time for both of the years' data. The refactored script used output arrays versus the original script used variables. The nature of using an array leads to a faster processing time by having the ability to store multiple values while the variable is only able to use one value at a time. The example from the refactored code is shown here:
```
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
```

A disadvantage to a refactored script is the need to understand the code well enough in the first place to be able to read through and know where to condense code. The ability to debug the large amounts of code involved in a rescript is important to the timely nature needed to finish a project for the client.

In conclusion, both codes worked well and gave the same output in data. The general public are fans of speed, so the client will be happier with the ability to process their data 85% faster.
