# Stock Analysis
This project is an analysis of green energy stocks using VBA to measure performance over the 2017-2018 timeframe. The analysis looked at 12 stocks in total and was meant to measure trading volumes and total returns in order to assist the climate in making more informed choices using past stock performance and trading volumes. The final product for this [Analysis](https://github.com/AsaHolley/Stock-Analysis/blob/main/Challenge_AH.xlsm) was produced using VBA to run quickly using a refactored code that allowed the client to perform the analysis and see color coded results at the click of a button. 

## Comparison of Stock Returns and Refractored Code

From the below analysis it is clear between 2017 and 2018 a shift in sentiment happened on green stocks, with the notable expectations of ENPH and RUN. Both ENPTH and RUN produced outsized returns in both 2017 and 2018 while green stocks on a whole performed poorly after having a healthy 2017. It is also notable that both ENPH and RUN had higher trading volumes on average then their peer group in both years, which could indicate higher interest in these stocks from the broader market.  

**2017 Stock Returns and volume**
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/2017%20Analysis.png)


**2018 Stock Returns and volume**
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/2018%20Analysis.png)



**Original Code Run Time**

2017
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/Non-refractored%20Code%202017.png) 


2018
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/Non-refractored%20Code%202018.png)

**Refractored Code Run Time**

2017
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/Refractored%20Code%202017.png)


2018
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/Refractored%20Code%202018.png)

As the below screenshots shot, the refactored code for the green stock analysis improved the speed of the analysis significantly. The reason for this improvement is because the refactored code uses less computing power for calculations due to the fact that we are using more effect loops that go through fewer iterations. 

**Original Code**
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/All%20Stocks%20Analysis%20VBA%20Code.png)

**Refractored Code**
![](https://github.com/AsaHolley/Stock-Analysis/blob/main/All%20Stock%20Analysis%20VBA%20Code%20Refractored.png) 

## Summary


**Advantages or Disadvantages of refactoring code?**
The purpose of refactoring code, and its primary advantage, is to make the code efficient. By being more efficient the code is able to run faster, be updated or edited in a quicker manner, and can be duplicated or reused with relative ease.  The major disadvantage the general code is more sequential than the refactored code and therefore can be easier for someone not familiar with the code to read.


**Pros and cons of applying refactoring to the green stocks VBA script?**
The major pro is the increased speed and efficiency with which the code is executed. Through refactoring less loops are needed and therefore the code runs faster. With less loops it is also easier to add news lines of code and modify the existing code. The major disadvantage, however, is that the refactored code is harder to read and understand if you are not the originator and take long breaks in between looking at the code. 


