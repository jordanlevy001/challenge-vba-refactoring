# Module 2 Challenge:
# Refactoring Code in VBA


## Results Analysis


### Project Overview

#### The purpose of this analysis is to refactor the module 2 solution code. The refactored code should loop through all of the stock data one time to collect information on each stock's total daily volume and the return. Additionally, we will determine the new analysis run time and evaluate how it compared to the run time of the subroutine before the code was refactored.


### Results

#### Almost all of the stocks we examined in Steve's workbook had a positive percent return in 2017. The only stock that had a negative percent return in 2017 was "TERP". But in 2018, almost all of the stocks had a negative percent return. There were 2 stocks, "ENPH" and "RUN", that had positive returns in 2018. Overall, the majority of the stocks had better performance in 2017 than 2018.

<img width="352" alt="All Stocks Analysis 2017" src="https://user-images.githubusercontent.com/88804543/131187934-cc2bd676-6af6-4a6d-bfdd-507971e014e2.png">
<img width="349" alt="All Stocks Analysis 2018" src="https://user-images.githubusercontent.com/88804543/131187939-dc51e216-1bb9-46d7-b944-23368ba3f5ab.png">

#### After refactoring the code, the run times were reduced by about 4X. The new run times were 0.0703125 seconds for both year 2017 and 2018. The previous run times from the module were about 0.28 seconds for both year 2017 and 2018.

<img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/88804543/131187960-89eeef9e-345c-476c-982b-1b7aeb1cd124.png">
<img width="254" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/88804543/131187979-ab31be6b-b52c-4213-9f6c-045924a5136b.png">

#### When we refactored the code, we created output arrays for each ticker's volume, starting price, and ending price. This improved our overall logic and allowed us to write the conditional statements more efficiently, utilizing a tickerIndex. 

<img width="672" alt="Output Arrays and Conditionals" src="https://user-images.githubusercontent.com/88804543/131192605-f57d34dc-d786-4ca8-be58-e7f2aad9aba8.png">


### Summary

#### Refactoring code is an important part of the coding process. Usually the first draft of code can be streamlined and improved upon. There are many advantages to refactoring code; these advantages include improving the logic, organization, and readability of the code. Also, removing bugs and vulnerabilities in the code. On the other hand, there are also disadvantages to refactoring code. Some disadvantages include that refactoring code can be time consuming or increase the chance of a mistake being made as the code is worked over multiple times.

#### For this project, refactoring the code improved the organization, readability, and run time. The refactored code will also help Steve in the future, when he needs to analyze multiple stocks and needs the analysis to run quickly. The disadvantage in this case was refactoring the code was somewhat time consuming.
