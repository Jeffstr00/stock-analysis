# VBA of Wall Street

## Overview

We are in charge of assisting Steve, a recent finance graduate, with analyzing green stocks in order to help his parents hopefully make a good investment.  Their current plan is to invest on a single stock solely based on the fact that its ticker symbol coincides with the location of their first date.  In order to see if that is a good idea or not, we looked at two years' worth of data on a dozen different green stocks, which covered thousands of data points.  Since we were dealing with such a large mountain of information, we used Excel and it's Visual Basic editor to run computations for us and then display useful information that can hopefully help Steve and his parents made a good financial decision.

## Results

### Calculating Each Stock's Return Using Starting and Ending Prices

In order to see how each stock performed each year, we ultimately compared its ending price to its starting point (we also counted the number of trades to see how active each stock was).  To find the starting price, we used the following formula: `If Cells(j - 1, 1).Value <> Cells(j, 1) Then tickerStartingPrices(tickerIndex) = Cells(j, 3).Value`.  This checked to see if we were on a new stock by seeing if the current ticker matches the former.  If so, we noted its starting price.  We used a similar formula to determine the ending price: `If Cells(j + 1, 1).Value <> Cells(j, 1) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value`.  After this, a simple `tickerIndex = tickerIndex + 1` increased the tickerIndex so that we could move on and do the same for the next stock.  Once we had starting and ending prices for each stock in a given year, we could determine the return using the following formula: `(tickerEndingPrices(k) - tickerStartingPrices(k)) / tickerStartingPrices(k)`.

### 2017 + 2018 Stock Returns

We ended up with the following results for each year:

![2017 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2017_orig.png) ![2018 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2018_orig.png)

While Steve is the financial expert and not us, it is crystal clear that, for whatever reason, these green stocks had excellent performance in 2017 as only one out of the twelve went down in value (and even then it was only a meager loss).  However, 2018 was the exact opposite, as only two stocks showed positive results.

## Summary
