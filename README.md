# Stock Analysis with VBA

## Overview of Project
Steve is helping his parents invest in the stock market. They want to invest in green energy stocks, specifically DAQO Energy Corporation (DQ). Steve wants them to diversify their account, so he has compiled information on several green energy company stocks and wants us to analyze the data to help him understand it. Once we have done this, we want to refactor our code to loop through all available data more quickly.

## Results and Analysis

### Analysis
The analysis method used was dependent on two main factors: there were only 12 stocks data and the data was organized alphabetically by ticker and choronologically within the ticker. First the macro asked the user for the year they wanted it to summarize and it established the necessary variables and arrays. Then it utilized a "for" loop to run through all the rows of data. Within that loop, an array was used to store the net volume of stock traded for each ticker, both the starting prices and the ending prices for each ticker were saved, and a variable to track the ticker for the row of data. After the loop was closed, the data was output onto an Excel worksheet and was formated for ease of understanding. The time it took for the macro to run was also output. The macro could be activated by pressing a button in the summary Excel sheet. All data could similarly be cleared between macro uses by pressing a different button. These buttons are shown in Figure 1.

![MacroButtons.png](/Resources/MacroButtons.png)

Figure 1. Screen shot of the buttons in the Excel sheet to run macros.

### Results
The macro produced a summary table that is shown in Figure 2 for 2017 and Figure 3 for 2018.

![VBA_Challenge_2017.png](/Resources/VBA_Challenge_2017.png)

Figure 2. Table of stock tickers, total daily volume, and the percent return for 2017 data.

![VBA_Challenge_2018.png](/Resources/VBA_Challenge_2018.png)

Figure 3. Table of stock tickers, total daily volume, and the percent return for 2018 data.

The message box produced by the macro to inform the user of how long it took to run the analysis was saved for 2017 and 2018, and are shown in Figure 4 and Figure 5 respectively.

![VBA_Challenge_2017_Results.png](/Resources/VBA_Challenge_2017_Results.png)

Figure 4. Screen shot of macro runtime length for 2017 data.

![VBA_Challenge_2018_Results.png](/Resources/VBA_Challenge_2018_Results.png)

Figure 5. Screen shot of macro runtime length for 2018 data.

Before the refactor, the macro took 0.8632813 seconds for 2017 data and 0.8554688 seconds for 2018 data, and a separate macro was needed to format the data.

## Summary

### Refactoring in General
Refactoring a macro can greatly increase the performance, can allow the macro to be applied to new data more easily, and can uncover parts of code that could cause errors. It can also make the macro easier to understand if someone else is trying to edit or use it. The disadvantage of refactoring a macro is the limiting returns of time saved versus time spent refactoring.  If the macro is not used for anything else in the future, the time spent to streamline the code will not be worth the improvement in speed. If the code will be used many times and the time saved by refactoring is noticeable, then it is worth it.

### Refactoring This Code
By refactoring this code, the macro was able to perform the same function and more (formats the output) in less than 15% of the time for 2017 data and less than 12% of the time for 2018 data. This is a significant improvement in the speed of the macro, and a few minor changes would enable this macro to include more stocks. A disadvantage of refactoring this code is the amount of time saved does not justify the amount of time spent refactoring. The macro was already running under one second and it could be formatted in one mouse click (since there was already a button to run the format macro). Spending 30 minutes to save 0.7 seconds on a macro that may only be run a couple of times is not worth the effort.