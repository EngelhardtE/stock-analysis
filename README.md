# stock-analysis

## Purpose
It may seem difficult to analyze stock prices due to prices changing throughout individual trading days and from day to day. This can lead to hundreds of price points in any given year. However, by using VBA you can pick out each stock's starting and ending price in order to see the return you would receive if you held it throughout a given year. You can also measure the stock's total daily volume. Using this information, you can make educated decisions on how to build your stock portfolio while also having code that can be reused for years to come. 

## Results

### Methodology
The general idea was to use a array and a nested loop to measure the returns using the starting and ending prices for each individual stock in the portfolio for the years of 2017 and 2018. The nested loop determined which prices were the starting and ending points for each stock in a given year, as well as the overall daily traded volume. The final loop would then use that information to organize the total traded volume and yearly return for each of the eleven stocks. This can be seen as follows:

    For i = 0 To 11
        ticker = tickers(i)
        tickerVolumes = 0
        Worksheets(yearValue).Activate
        For j = 2 To RowCount
        If Cells(j, 1).Value = ticker Then
              totalVolume = totalVolume + Cells(j, 8).Value
            End If
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                startingPrice = Cells(j, 6).Value
            End If      
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1) = ticker Then
                endingPrice = Cells(j, 6).Value
            End If
        Next j     
        Worksheets("All Stocks Analysis").Activate       
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1    
    Next i
 
Once the loop was completed, a few more lines of code were added to format the information in a visually pleasing manner. An additional loop was also written to color code the returns (green for positive, red for negative). 

    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i

### 2017
For the portfolio of eleven given stocks, 2017 was a year of positive trends. All stocks except for TERP saw positive returns, with DQ nearly hitting a 200% return rate. It should be noted that most stocks in the portfolio did not see such high returns. Some stocks, such as RUN and AY were much lower, although still positive at 5.5% and 8.9% respectively. TERP saw a loss of 7.2%, but that would be counteracted by the returns of the other stocks, assuming that the portfolio was evenly balanced between all eleven. 

![2017data](https://imgur.com/fy5njfJ)

### 2018
2017 may have been a great success for the portfolio, but 2018 saw losses almost across the board. All of the stocks saw losses, except for ENPH and RUN. In fact, while RUN saw the lowest positive return in 2017, it saw the highest return in 2018 at 84%, with ENPH slightly behind at 81.9%. DQ went from the highest return to a loss of 62.6%, the lowest of the entire portfolio. Regarding total daily volume, 2018 saw decreases for the majority of stocks. The more sparsely traded stocks (AY, CSIQ, and DQ) all saw slight increases, while the rest decreased.  

![2018 data](https://imgur.com/c3V6EdT)

## Summary

### Refactoring Code in General
There are a swath of advantages and disadvantages to refactoring your code. The obvious advantage is that it allows the code to run faster and smoother. While an extra second or fraction of a second may not seem like much at first, it can make a big difference if the code is being ran hundreds or thousands of times throughout its span of use. For every 1,000 times code is used, refactoring it to save 0.1 seconds saves the user 1 minute 40 seconds of time. However, if not done carefully refactoring code can also have dangerous disadvantages. Firstly, if done sloppily the code can lose some parts of its functionality. In cutting corners, one could accidently stop the code from accomplishing a certain part of its purpose. Additionally, refactoring code could make it harder for others to interpret the code.

### Refactoring this Specific VBA Script
Once again, the advantage of this specific VBA script is that it runs slightly faster. While the extra tenth of a second may not seem like much, if it were to be used continuously for years on end to analyze stocks then the time saved would add up. However, I do believe a stark disadvantage is that it may be difficult for somebody understand why it was written the way it was. At the very least, they would not get the complete picture that I saw when writing it, even if they do fully comprehend what each line does functionally. 
