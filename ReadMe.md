# Green Energy Ticker Analysis

## Overview
The purpose of this project is that Steve's parents wanted to invest into renewable energy companies but do not know a lot about that field. So they had invest into one company, DAQO New Energy Corp, because it had sentimental reasons for them.  Steve thought it best to instead analyze other renewable energy companies as well to help diversify his parent's portfolio and has enlisted us to help. We are here to make a simple script for him to give a quick analysis that his parents can see and understand.

## Results of Tickers
From the data we have gathered the stocks compared from 2017 ![StockPerformance2017.PNG](\Resources\StockPerformance2017) and 2018 [!StockPerformance2018](Module 2\Resources\StockPerformance2018) show that the companies that Steve has found won't perform well for the most part. I had combed through the data and checked the volume using this line: 
```
If Cells(j, 1).Value = tickers(tickerindex) Then
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(j, 8).Value
        End If
```
which gives me the volume to post into the analysis. The tickerindex variable being which ticker we are currently running through in this dataset.  Then after the volume was calculated I had grabbed the beggining and end price of the stock:
     
```
If Cells(j - 1, 1).Value <> tickers(tickerindex) And Cells(j, 1).Value = tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(j, 6).Value
        End If
            
If Cells(j + 1, 1).Value <> tickers(tickerindex) And Cells(j, 1).Value = tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(j, 6).Value
        End If
        
        Next j
```
and this gave me the ending price and then the begginning price so Steve can see how it changed over time.  After running this code originally the times were quick being at 0.6406 seconds and 0.6483 seconds respectfully for 2017 [!VBA_Challenge_2017_Unfactored](Module 2\Resources\VBA_Challenge_2017_Unfactored) and 2018[!VBA_Challenge_2018_Unfactored](Module 2\Resources\VBA_Challenge_2018_Unfactored) respectfully. After some minor cleaning of the code I had gotten 2017 down to 0.2304 seconds [!VBA_Challenge_2017](Module 2\Resources\VBA_Challenge_2017) but ran into issues with 2018 not giving me back a response, and ultimately no longer giving me a result for 2017 which proved to be troubling.

##Results
From a glance of the two pictures of the 2017 and 2018 performance its clear that the stocks that Steve chose, apart from two, had proven to do poorly in 2018 compared to 2017. ENPH and RUN had another great year of returns compared to the rest on this table. It's clear that the stocks he chose have to be dwindled down to these two only as safe investments for his parents.  The last thing I can infer on is that even if the yearly return of these 12 stocks had only 20% of them be in the green, the volume could be a good sign if one of his parents were a daytrader. That being that daily volume being high is a good sign that puts and shorts are being made but that isn't investing.  

## Summary of Data
What is interesting with code is at a novice level it will be more complex than to say a professional, who would have more knowledge of cleaning and simplifying scripts for quicker use deployment.  With the idea of refactoring, that's what you get, the code can be simplified and cleaned for a better human to human read that in the case of future updates it can be easily manipulated and updated with a greater convenience for the developer. Along with that, the script will run faster because the compilier has less to read as it runs through the script. Even with refactoring, it could be a waste of time depending on what is being refactored. I do not believe small datasets benefit from this kind of cleaning because the scripts are barely noticable before and after being refactored. 
### Refacorting Opinion
Which is where I get into the refactoring of our VBA code. I felt that doing this refacoring was a waste of time (In the sense of an analysis for Steve, I'm still learning I think it is important). The idea of the script was to save Steve about 10 or 15 minutes of work by automating the worksheet to output the data for him. The original script though however was running about at .6 seconds for me, so I saved .4 seconds when refactoring it which is not a lot of time compared to the multiples of minutes saved from doing the original. If the data was ten times this current amount, it would make sense to refactor only for the reason that the script would take a good amount longer being the data to sift through is much more substantial.


