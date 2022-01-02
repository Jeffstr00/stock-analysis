# VBA of Wall Street

## Overview

We are in charge of assisting Steve, a recent finance graduate, with analyzing green stocks in order to help his parents hopefully make a good investment.  Their current plan is to invest in a single stock simply based on the fact that they like helping the environment and its ticker symbol coincides with the location of their first date.  In order to see if that is a good idea or not, we looked at two years' worth of data on a dozen different green stocks, which covered thousands of data points.  Since we were dealing with such a large mountain of information, we used Excel and it's Visual Basic editor to run computations for us and then display useful information that can hopefully help Steve and his parents made a good financial decision.

## Results

### Calculating Each Stock's Return Using Starting and Ending Prices

In order to see how each stock performed each year, we ultimately compared its ending price to its starting point (we also counted the number of trades to see how active each stock was).  To find the starting price, we used the following formula: `If Cells(j - 1, 1).Value <> Cells(j, 1) Then tickerStartingPrices(tickerIndex) = Cells(j, 3).Value`.  This checked to see if we were on a new stock by seeing if the current ticker matches the former.  If so, we noted its starting price.  We used a similar formula to determine the ending price: `If Cells(j + 1, 1).Value <> Cells(j, 1) Then tickerEndingPrices(tickerIndex) = Cells(j, 6).Value`.  After this, a simple `tickerIndex = tickerIndex + 1` increased the tickerIndex so that we could move on and do the same for the next stock.  Once we had starting and ending prices for each stock in a given year, we could determine the return using the following formula: `(tickerEndingPrices(k) - tickerStartingPrices(k)) / tickerStartingPrices(k)`.

### 2017 + 2018 Stock Returns

We ended up with the following results for each year:

![2017 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2017_orig.png) ![2018 Stock Results](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2018_orig.png)

While Steve is the financial expert and not us, it is crystal clear that, for whatever reason, these green stocks had excellent performance in 2017.  In fact, only one out of the twelve went down in value (and even then it was only a meager loss).  However, 2018 was the exact opposite, as only two stocks showed positive results.  This shows how putting all of your eggs in one basket (in just green stocks, especially just in one individual stock, or even in the stock market in general) can be risky.  While 2017 was a good year, if they went with their plan of investing in DQ and bought in at the beginning of 2018, they would have lost 62% of their investment, which is catastrophic!  For DQ in particular, their relatively low volume in 2017 (other stocks trade roughly 3-20x more often) could indicate that it's a newer, more volatile stock.

### Refactored Script and Execution Times

In order to measure the performance of our coding, we included a display to show how long it took to run the program.  Both 2017 and 2018 took roughly 2/3rds of one second.  While we certainly don't think that would be too long for Steve and his parents to wait, we wanted to see if we could make things more efficient in case we wanted to reuse this code to look at substantially larger amounts of data (or if we were in a situation where fractions of a second did matter, such as making real-time, high-volume stock trades).  Instead of the original code where we went through each row with the following If And Then statement: `If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then startingPrice = Cells(j, 6).Value`, we refactored the code and set it up so that we set up arrays to keep track of the starting price, ending price, and traded volume for each ticker.  We then only bothered to change the starting and ending prices when we reached the beginning or end of a new group of tickers, making things much more efficient.  While the results and the output did not change, performance did as times were slashed from 0.668 and 0.660 seconds to 0.0820 and 0.0977 seconds for 2017 and 2018 respectively, making it approxiately 10 times faster!  Again, it's doubtful that this increase in speed will be appreciated by Steve and his parents, but if the code were to be used in more demanding situations, the difference could be important.

![2017 Stock Results Refactored](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2017.png) ![2018 Stock Results Refactored](https://github.com/Jeffstr00/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary

### Refactoring in General

While the refactoring of the code did not change the results or the output one bit, it is still an important practice to use in coding.  Refactoring is when you go back and make code either more efficient or even just easier to follow.  Making code run more efficiently can end up saving time, resources, and ultimately money in the long run.  Even slight improvements can be magnified and add up when they are repeated possibly millions of times.  Making code easier to read, understand, and follow can pay dividends either when you have to go back and work with your code again or especially when other people work with your code.  However, refactoring isn't completely without its downsides.  It can be rather time consuming, so if you don't end up working with this code again, that time was pretty much wasted.  For instance, if we refactored this code simply to save Steve and his parents a fraction of a second, it would be hard to argue that was a good use of time.  It's also possible that you could accidentally break code (and not realize it) when making changes.  That's why it is important to go back and check to make sure that your code still runs as intended, even if the inputs change.

### Refactoring in this Case

In this case, there doesn't seem to be much upside to the original code, aside from it maybe being easier to follow on a step-by-step basis, since you are doing things individually.  The updated code is definitely more efficient, as instead of three seperate If Then statements (two of which are If And Then), it is condensed into just two simpler If Then statements where it checks to see if the current ticker is different from the one before or after it and acts accordingly.  This is evidenced by the fact that the new code ran approximately 10x faster.  Again, while this improvement in this example is insignificant on its own (and would clearly not be worth the investment in time it took), that entirely small (but proportionally big) improvement could be very beneficial if this code were to be reused in other more demanding situations.
