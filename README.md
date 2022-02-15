# Module 2 |  Wall Street Assignment

Explore green energy stock performance by analyzing financial data using VBA.

## Overview of Project
### Purpose
The purpose of this analysis is to gain insights on the stock performance of various green energy stocks through their total volume of trade and yearly growth in VBA.  Presumably this is to help Steve's parents better select stocks in green energy, as they were entirely invested in the stock DQ.  Analysis is done by year offering a chance to compare.  This is also an opportunity to refactor the original code used for this analysis.

## Analysis and Challenges

### Analysis of Green Stocks
This analysis used VBA to determine total volume and yearly growth.  A ticker was used to keep track of a given stock.  To obtain yearly growth, the volume traded in a given day was added to a cumulative total.
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/volume.png)

To obtain the yearly return, the first and last instance of a stock return was saved into a variable, which were then used to calculate the yearly return.
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/return.png)

Finally, the results were output and formatted to easily identify returns that increased or decreased in their respective year.  A timer was also included to measure the speed of the code.

### Refactoring
A second version of the program was created to try to increase the speed of the analysis.  The primary change occurred in how the program iterated through the stocks.  In the first case, the program went through each row for each ticker.  The second case instead used a ticker index to keep track of the given stock, and change once the stock was no longer appearing.  This allowed for the program to run through the rows one time.


### Challenges and Difficulties Encountered
While using the static ticker to record the outputs for each stock, I also left a static ticker in the outputs which lead to not outputs.  I realized I had to run through each ticker in the output in order to show the outputs.
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/mistake.png)

## Results

- *Comparing 2017 to 2018 Volume and Return*
Taking a bigger view of results from 2017 and 2018, its clear green energy had a good year in 2017 and a bad year in 2018.  If one were to insist on staying in green stocks, it may be the case that you want to look into minimizing hard.  Taking a specific look at the DQ stock still gives us reason to worry.  In 2017 DQ had a good year, but was not traded significantly.  The following year it traded more which arguably makes it more accurate, and had among the worst results of 2018.  By contrast ENPH had a decent year in 2017 and strong year comparatively in 2018, while trading at high volumes.  

2017
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/results_2017.png)
2018
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/results_2018.png)
- *Refactoring*
The execution times between the original and refactored code were decisive, with the refactored code moving roughly 10 times faster as compared with year 2018 seen here.

Original Code:
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/Readme_images/old_code.png)
Refactored Code:
![enter image description here](https://raw.githubusercontent.com/Andres-T-Hernandez/stock-analysis/main/Resources/VBA_Challenge_2018.png)

## Summary
### Advantages and Disadvantages to Refactoring
When considering refactoring, the primary disadvantage to refactoring is of course the extra time it takes to redo a completely functional program.  If the program is bigger, it may take significant time and could mess with other parts of the program, requiring lots of testing.

Of course the benefits are a faster program able to work with more data.  The code itself is also cleaner and easier to follow it seems.

### Advantages and Disadvantages to the Wall Street Assignment VBA refactoring
The primary advantage to the original code is in its intuitive algorithm.  When planning out how to write the code, it is easy enough to run through the rows of data with each ticker using a for loop.  This gets the program up, running, and functional when there is a lot of initial planning.  However reading the code itself is messier, and on a second thought it should seem clear running through each row multiple times is not optimal.  Refactoring then creates cleaner code, and although less intuitive from a planning perspective, since you have to consider the extra ticker index variable, once the program is established its easy to focus on specific ways to make the improvement making it more efficient and easier to follow when making further adjustments.

