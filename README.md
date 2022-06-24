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


### Summary of Refactoring Code


    Their are both advantages and disadvantages of refactoring code. The advantage of refactoring code is rewriting code that runs more efficiently and quicker. Complicated code that require redundant or repititous calling can prove to be expensive in resources, especially if the code is iterating over thousands or millions of lines. While there was a slight difference in the refactoring I displayed above, the script was not substantial nature, either. However, if the code were hundreds of lines, refactoring would have proven to be beneficial and less taxing on resources running the script. 

    The disadvantage to refacorting code is that it might take some time to figure out the logic on how to best refactor the code. Sometimes the time it takes to figure out the logic is much more than the difference in time it would save if you didn't refactor the code in the first place.

    These pros and cons apply in the sense that while I did refactor the code to run more efficiently, in the grand scheme of things, the time and resources saved is insignificant. The data set for 2017 had the biggest change in time when compared between original script and refactored. The difference the refactoring did here was ~1.4 seconds. There was a slight, if any change for the year 2018 between the original script and the refactored script. All in all, despite the pros and cons of refactoring, one should always build a script to run in the most efficient manner. This will ensure that it runs efficiently everytime, and will help the next person after you who might need to work and understand the script as well. 