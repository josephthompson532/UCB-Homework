# VBA Stock Anaysis Automation

## Overview of the Project
In this project we automated an analysis of stocks using VBA and then refactored the program to make it run more efficiently. While there are many ways to write a program, not all programs run equally quick. Through rethinking the code we were abe to make a quicker program that calculated the Total Daily Colume and Return for any given ticker on a spreadsheet. We then automated the formating of the output data as well and finished off by adding easy to use Buttons for the end user.

### Purpose
The purpose of this project was to show how changes to the stucture of code can cause the program to be more efficient by a factor of many times. In this particular circumstance, the presence of a nested loop made the program much slower than it need to be. 

## Analysis and Challenges of Refactoring
This challenge was fairly straightforward to analyze and I encountered very few challenges while attempting it. I would however not have known to make these adjustments to make the code run faster. I will need more experience to have a better feel for which design patterns will be more efficient than others.

### Analysis of Refactoring
The nested loop was the logic really holding the program back. It takes much more time to do a nested loop than it does to simply index every variable at a desired tickerIndex. These constitute two distinct coding design patterns. This subtle change allowed the program to loop through each ticker and get the starting and ending prices as well as summing the ticker volume without running inside another loop. I can see how this woul be faster. Ultimately, it was faster by a factor of 5 when we ran the program using a nested loop and the Timer function.

### Challenges of Refactoring
A particular challenge for me was grasping the scope of what I needed to do with regards to the 
news setup for the program. For instance, had I been coding alone, I likely would have forgotten that I would need to create an array for TickerVolume as well as tickerStartingPrices and tickerEndingPrices. I also had a hard time keeping track of which variables I was using for the for loops and whether or not they were allowed to be the same variable.

## Results
The initial program that used the for loop ran in .5117 seconds while the refactored program ran in .10156 seconds. This means my refactored code was many times faster.
