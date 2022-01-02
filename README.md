# VBA of Wall Street

## Overview

We are in charge of assisting Steve, a recent finance graduate, with analyzing green stocks in order to help his parents hopefully make a good investment.  Their current plan is to invest on a single stock solely based on the fact that its ticker symbol coincides with the location of their first date.  In order to see if that is a good idea or not, we looked at two years' worth of data on a dozen different green stocks, which covered thousands of data points.  Since we were dealing with such a large mountain of information, we used Excel and it's Visual Basic editor to run computations for us and then display useful information that can hopefully help Steve and his parents made a good financial decision.

## Results

### Calculating Each Stock's Return Using Starting and Ending Prices

In order to see how each stock performed each year, we ultimately compared its ending price to its starting point (we also counted the number of trades to see how active each stock was).  To find the starting price, we used the following formula: `If Cells(j - 1, 1).Value <> Cells(j, 1) Then tickerStartingPrices(tickerIndex) = Cells(j, 3).Value`.  This checked to see if we were on a new stock by seeing if the current ticker matches the former.  If so, we noted its starting price.  We used a similar formula to determine the ending price: `If Cells(j + 1, 1).Value <> Cells(j, 1) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value`.  After this, a simple `tickerIndex = tickerIndex + 1` increased the tickerIndex so that we could move on and do the same for the next stock.  Once we had starting and ending prices for each stock in a given year, we could determine the return using the following formula: `(tickerEndingPrices(k) - tickerStartingPrices(k)) / tickerStartingPrices(k)`.

### 2017 + 2018 Stock Returns

We ended up with the following results for each year:

![2017 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2017_orig.png) ![2018 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2018_orig.png)

While Steve is the financial expert and not us, it is crystal clear that, for whatever reason, these green stocks had excellent performance in 2017.  In fact, only one out of the twelve went down in value (and even then it was only a meager loss).  However, 2018 was the exact opposite, as only two stocks showed positive results.  This shows how putting all of your eggs in one basket (in just green stocks, especially just in one individual stock, or even in the stock market in general) can be risky.  While 2017 was a good year, if they went with their plan of investing in DQ and bought in at the beginning of 2018, they would have lost 62% of their investment, which is catastrophic!  For DQ in particular, their relatively low volume in 2017 (other stocks trade roughly 3-20x more often) could indicate that it's a newer, more volatile stock.

### Refactored Script and Execution Times

In order to measure the performance of our coding, we included a display to show how long it took to run the program.  Both 2017 and 2018 took roughly 2/3rds of one second.  While we certainly don't think that would be too long for Steve and his parents to wait, we wanted to see if we could make things more efficient in case we wanted to reuse this code to look at substantially larger amounts of data (or if we were in a situation where fractions of a second did matter, such as making real-time, high-volume stock trades).  Instead of the original code where we went through each row with the following If And Then statement: `If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then startingPrice = Cells(j, 6).Value`, we refactored the code and set it up so that we set up arrays to keep track of the starting price, ending price, and traded volume for each ticker.  We then only bothered to change the starting and ending prices when we reached the beginning or end of a new group of tickers, making things much more efficient.  While the results and the output did not change, performance did as times were 

![2017 Stock Results Refactored](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2017.png) ![2018 Stock Results Refactored](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary
