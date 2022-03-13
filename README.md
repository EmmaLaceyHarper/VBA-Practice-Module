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

![Ticker Index](https://user-images.githubusercontent.com/99626046/158083644-453116a5-4262-4f0c-92b0-597cf4f2c4a6.PNG)

From there via a series of if statements the code loops through the data and brings in the total volumes and returns for the years for each indexed stock. 

![Loops](https://user-images.githubusercontent.com/99626046/158083674-f2b9bcf6-1dad-4a29-b847-d9d7fca48f94.PNG)

Through our creation of the stock index we were able to analyze not only one stock's data but all twelve stocks with volume and return data on the 2017 and 2018 tabs. We then created code that formatted the data to make it much more readable with the negative returns highlighted in red and positive returns highlighted in green. This made Steve's job much easier in terms of investment analysis. 

![Formatting Code](https://user-images.githubusercontent.com/99626046/158083677-74560594-6a25-4b07-a8c9-0e263bbbe853.PNG)
