# VBA of Wall Street

## Overview of Project
Steve wanted to help his parents diversify their green energy portfolio.  He asked us to use VBA to produce a report of the Total Volume and the Annual Return of several companies focused on renewable energy.

## Results

### Stock Performance
2017 was a great year to be in the renewable energy sector.  All but one of the companies from our analysis experienced postive annual returns.  The average return for 2017 was 67.3%!  Four companies more than doubled their starting price and DQ (the company Steve's parents have invested their money) was .6% away from tripling it.  2018 was a different story.  With only two companies posting positive gains, the average annual return was -8.5%. The industry seems to be volatile to say the least.  Fortunately, with the analysis we have done, we can advise Steve's parents to diversify their portfolio an spread their risk.  While past performance is not always an indicator of future success, I'd suggest to invest some of their money into ENPH and RUN, the only two companies to post positive annual returns both years. We calculated annual returns using `if-then` statements and `for` loops as in the following block of code: 
     
        ```
       '3b) Check if the current row is the first row with the selected tickerIndex.
         
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
         
         tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
         
         End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
         
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If'
        ```
