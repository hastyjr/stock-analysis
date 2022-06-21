# :chart: stock-analysis

## :open_book: Overview of Project
### Steve's Stock Analysis :chart:
 _The purpose of this project is to analyze Steve's provided stock dataset. His intent for this project was to go deeper into the data to present to his parents. This analysis will provide what the outcome of the dataset is in summary and provide insight on how much time Steve can save now that we have automated some of the processes through VBA scripting and macro creations._

---
## :part_alternation_mark:	 Results

###  Optimizing Code with Refactoring
_As the orignal script was written you will see how the code was written in a way that can fetch the data in a timely manner, but not ncessarily efficient as possible as we compare both years.

When attempting to optimize the code to run more efficiently by refactoring, we implmented several steps to get there._

1. Created a `tickerIndex` with three output arrays: 
    * `tickerVolumes`, 
    * `tickerStartingPrices`, and 
    * `tickerEndingPrices`

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/code/1.png)

2. Created a `for` loop
    * initialized the `tickerVolumes` to `zero`
    * looped over all rows in the spreadsheet

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/code/2.png)

3. Wrote an inside `for` loop to
    * optimize iterating over the `tickerIndex` variable in order to increase the current `tickerVolumes`
    * Wrote IFTTT statements to check if current/last rows and assign either `ticketStartingPrices` or `tickerEndingPrices` output.

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/code/3.png)

4. Used the `for` loop to loop **through** the arrays

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/code/4.png)

### A Comparison of Execution Times


#### _**2017 Original Script vs Refactored**_

:chart_with_downwards_trend: _**2017 Original Script**_

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/2017%20-%20original%20script.png) 


:chart_with_upwards_trend: _**2017 Refactored**_

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)


#### **2018 Original Script vs Refactored**

:chart_with_downwards_trend: _**2018 Original Script**_

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/2018%20-%20orignial%20script.png)
     
:chart_with_upwards_trend:	 _**2018 Refactored**_

![This is an image](https://github.com/hastyjr/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
     

--- 
## :closed_book: Summary
 In a summary statement, address the following questions.

### Summary of Refactoring Code
* What are the advantages or disadvantages of refactoring code?

* How do these pros and cons apply to refactoring the original VBA script?
