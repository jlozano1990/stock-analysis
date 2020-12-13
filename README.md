# Stock-Analysis with VBA
## Overview of Project
### Purpose
The purpose of this challenge was to see if refactoring the VBA code we used in our first analysis would make it more efficient, so as to be able to use it on bigger datasets in the future.

## Results
### Stock Analysis
Simply by looking at the results of the analysis on the stocks in 2017 as compared to 2018, we can see the stock market return was significantly better in 2017 as opposed to 2018.

#### Stock Analysis Comparison
![Stock Analysis 2017](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/stock%20analysis%202017.PNG)
![Stock Analysis 2018](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/stock%20analysis%202018.PNG)

### Difference between original code and refactored code
When running the analysis on both datasets, the original code and the refactored code both produce the same results on the datasets. However, the refactored code seems to produce the results in a  more efficient manner. The refactored code was significantly faster(and cleaner) than the original code.

#### Original Code Time
![Original Code](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/All_stocks_analysis_2017.PNG)
![Original Code](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/All_stocks_analysis_2018.PNG)

#### Refactored Code Time
![Refactored Code](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
![Refactored Code](https://github.com/jlozano1990/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)

The refactored code runs about 0.5 seconds faster in both cases, making it more efficient.

## Summary
### Advantages and Disadvantages of refactoring code
The biggest advantage of refactoring code is making it more efficient. Refactoring code correctly makes it run better and allows you to use it in other processes, not just the one you're working on at the moment. With the refactored code we created in this challenge, we could input more data and we'd only have to change a couple of variables in our code to make it work. If we input more data and try to use the original code to analyze it, there would be more variables we would have go in and change one by one. By refactoring code, we're preparing ourselves in case the dataset changes in the future, so we won't have to go in and change a lot of the code.

A disadvantage of refactoring code is the possibility of doing something incorrectly and having the code give you errors. That's why it is important to keep saving changes we make to our code, so in case something goes wrong, we have previous code that worked to revert to.

### Pros and Cons of refactoring original VBA script
When it comes to refactoring the original VBA script we had, refactoring the code made it more efficient in our case. The code ran noticeably faster and maintained the same accuracy when it came to the results produced.

As far as any cons there might be when it comes to refactoring the VBA script we had, the only con I could think of is that it takes more time to do. However, that is offset by the fact that our refactored code will be easier to update should the dataset change. I did have a couple of issues with the code not running correctly at first, but once I figured out where the errors were, I was able to resolve it and it worked fine.
