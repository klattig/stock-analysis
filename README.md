# stock-analysis

# Overview of Project
This project can be used to easily determine yearly stock Volume and Return Percentages for various Stocks on a large Data Sheet. Data for other stocks, as well as other years may be added for additional calculation. 
## Purpose
The Purpose of the project was to edit the existing code to compile various Marcos simultaneously. We attempted to loop through all the data one time in order to collect the information then we determined whether the refactored code successfully made the VBA script run faster. 
### Analysis and Challenges
We compared the time to run the Analysis of the Yearly Volume and Return Rates for 2017 and 2018. The client (DQ) was interested in comparing its Yearly Volume to competitors with a script that will run will more efficiently.
## Analysis of 2017
###   Old
  ![](https://github.com/klattig/stock-analysis/blob/main/Resources/2017%20Old%20Code.png)
#### New  
  ![](https://github.com/klattig/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
  
From the images above it can be seen the new code for the 2017 Analysis ran faster

## Analysis of 2018
####  Old
   ![](https://github.com/klattig/stock-analysis/blob/main/Resources/2018%20Old%20Code.png)
#### New
  ![](https://github.com/klattig/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

From the Images it can be seen the new code for the 2018 Analysis also ran faster
Challenges and Difficulties Encountered
The challenge in executing this script in a single loop lied was accomplished by creating an Index variable for the various ticker codes for each company. This “tickerIndex” Variable allowed us to write a single loop that would execute or both the Daily Volumes as well as the Year Return Rate. Below is an example of how the code was changed to reflect the new variable. 
#### Old
    For i = 0 To 11
        ticker = tickers(i)
        totalVolume = 0
#### New
     For tickerIndex = 0 To 11
            ticker = tickers(tickerIndex)
            tickerVolumes(tickerIndex) = 0
 
## Results
The code proved to be faster when run for both existing data sets (2017 and 2018). Table Formatting, Labeling, and Coloring or Return Cells were also added to the single Macro for further efficiency. By combining the macros, we both consolidated steps for a more efficient analysis, but we also decreased the time for the script to run. This code can only diagnose historic data and does not have any predictive modeling abilities.
## Summary

### Advantages and Disadvantages of the refactored code
In general, the advantage of the refactored code is performing the complete analysis faster, and in less steps. The cleaner code does however become more complex, and time was spent debugging changes. 

### Advantages and Disadvantages of the original refactored code
The original code may provide an advantage when it comes to breaking up the step of the codes in order to test new scripts, while the newly refactored code may require more debugging in order to test.  Both codes provide a fast calculation and output to a table to visually display the results. 

