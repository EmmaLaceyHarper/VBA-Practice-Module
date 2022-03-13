# Stock Analysis Refactor Challenge

## Overview
### Background 
In our Stocks Analysis project we are helping Steve analyze a set of Green Energy stocks for his parents. His parents currently own DAQO New Energy Corp (Ticker: DQ), but Steve is concerned about the diversification of their portfolio and therefore wants to analyze a handful of Green Energy stocks in addition to DQ. 
### Purpose
For this project we utilized VBA in excel in order to create macros that would not only help Steve analyze the current data set for his parents, but also create a macro that would be reusable for other sets of data. Through our project we created code that pulled in the Tickers for each stock as well as their Total Daily Volume and Return in for both 2017 and 2018 on to one page. This should help Steve look at these tickers on one page and make more informed decisions on their volume and performance, both crucial to understanding the future trajectory of the return of these stocks in the future. After creating these codes we also created functions to time the code to check for inefficiencies. Through this project it was determined that the code could be refactored to create a macro that can handle a larger data set and run faster and more efficiently.

## Results Overview
### Investment Results
The overall results of the project were mixed. In general Steve's idea about diversification seemed to be correct as DQ performed satisfactorily in 2017 but severly underperformed in 2018. In general, the Green Energy stocks analyzed by Steve had a good return and trading volume year in 2017 with the average trading volume sitting at 263,886,592 and the average return at 67.3%. DQ's return for that year was way above the average sitting at 199.4% yearly return, however the trading volume for DQ was lower than the average at 35,796,200 leading to questions about interest and liquidity of the stock in compairson to others in the peer group. Turning the page to 2018 the entire peer group suffered with the exception of just two names that had a postitve returns for the year and above average trading volumes, ENPH and RUN. Other than those two stocks the peer group saw negative returns and modest trading volume increases. The average return for the group in 2018 was -8.5% and the average trading volume was 275,503,813. 2018 was a rough year for DQ who saw a negative yearly return of -62.6% which was below the average and actually was the lowest returning stock in the peer group. Although trading volumes were higher for this year for DQ at 107,873,900 it was still below the peer group average. Please refer to the return and total daily volume return streams for 2017 and 2018 for further information on this data. Overall Steve should continue to monitor the peer group as a two year return window is considered short in the investing world, but should likely decrease his family's position in DQ when financially advantageous and buy positions in either ENPH or RUN. 

![Stock_Returns_2017](https://user-images.githubusercontent.com/99626046/158082560-85c19dff-5437-4573-88e2-225d4c44807f.PNG)
![Stock_Returns_2018](https://user-images.githubusercontent.com/99626046/158082567-9295d1e9-ec4c-4597-8d7f-45fb3d72f5f1.PNG). 

### Macro Results
As far as our contribution to Steve, I feel we were able to create two usable macros, one that runs significantly faster than the other. For our first macro 'All Stocks Analysis' we built code that was able to sort through two years worth of data and properly identify and total the various daily volumes and returns for each year separately. The code first asks for which year the Steve would like to look like on the All Stocks Analysis tab with a message box.

![Input Box Code](https://user-images.githubusercontent.com/99626046/158083611-73d5d0f7-aec3-4249-bc18-c69383df4d04.PNG)

The code then indexes the list of tickers that are on both pages of the excel document to create a map for the macro. 

![Ticker Index](https://user-images.githubusercontent.com/99626046/158084242-ce05cc04-9f2f-40e7-838a-bbb76c4dfe1f.PNG)

From there via a series of if statements the code loops through the data and brings in the total volumes and returns for the years for each indexed stock. And finishes a code for a timer that gives us the time in terms of seconds it takes to run this code. 

![Loops](https://user-images.githubusercontent.com/99626046/158083806-ab722454-3a58-4770-85e1-d188a65eb1f4.PNG)

Through our creation of the stock index we were able to analyze not only one stock's data but all twelve stocks with volume and return data on the 2017 and 2018 tabs. We then created code that formatted the data to make it much more readable with the negative returns highlighted in red and positive returns highlighted in green. This made Steve's job much easier in terms of investment analysis. 

![Formatting Code](https://user-images.githubusercontent.com/99626046/158083677-74560594-6a25-4b07-a8c9-0e263bbbe853.PNG)

### Refactored Macro Results
In order to help Steve further we created a refactored code to run the code more efficiently and have the ability to run on larger sets of data. This refactored code took much less time to run and helped us to bring in the same information in a more efficient manner. The original code took about .7 to .9 seconds to run for both years, please refer to the images below. 

![All_Stocks_Original_Run_Time_2017](https://user-images.githubusercontent.com/99626046/158083979-bd41174e-3835-4f16-ae09-f0b91e9b6551.PNG)
![All_Stocks_Original_Run_Time_2018](https://user-images.githubusercontent.com/99626046/158083988-0e108f25-ba2b-473f-b638-20d05a89ccd2.PNG)

Although the time to run these codes were pretty good, the refactored code took much less time about .08 to .1 seconds to run. Please refer to the images below. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/99626046/158084020-229c9f45-b993-422a-8281-5e9a49b6bca8.PNG)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/99626046/158084027-8b30c5c0-98c8-4e17-8b4d-634dc2bcf435.PNG)

The way we made this code run faster was through a few key changes to the code. We started in much the same way with our input box and indexed tickers, but then we created a ticker index variable and three output arrays tickerVolumes, tickerStartingPrices and TickerEndingPrices to run through our if statement loops. 

![Ticker Index Variable](https://user-images.githubusercontent.com/99626046/158084341-d3e88301-fb80-467c-800d-37d740f0559b.PNG)

We then ustilized these arrays within the for loops and utilized the tickerIndex variable to move through the ticker index array each time the code ran through the loops. 

![Loops Refactored](https://user-images.githubusercontent.com/99626046/158084505-f3914af3-1434-4553-82ad-27c016adaf83.PNG)

Overall this made the code much faster and will helpfully continue to help Steve in the future with various projects. He would essentially only need to edit the ticker index array to reflect the tickers he was looking to analyze to run this code on a similar set of data with different or more tickers. 

## Summary
Overall the refactoring helped to decrease the time for the code to run and will help Steve to run larger data sets with the same code edited to reflect a new ticker array. Overall I think there were mostly advantages to refactoring the code, however one disadvantage might be the increased complication of the code to those that are less familiar with VBA. I think although the refactored code is more efficient it is significantly harder for the less VBA inclined reader to sort through. The original code was very straight forward and ran through the information withougt having to write in more variables and arrays, however once one understand the refactored code it can be used on larger sets of data by editing the ticker index array which is clearly the biggest advantage of the original versus the refactored code. 
