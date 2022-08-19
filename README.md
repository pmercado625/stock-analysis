# Stock Analysis Program
## Overview of Project
This project was created in order to assist in analyzing a dataset for stocks. By taking data from each stock ticker, this project can provide insight into various facets of the stock, such as the returns and the total volume traded. The project was then rehashed inorder to refactor the code into a more efficient process and then comparing the original with the refactored code.

## Results
![2017 stocks results](https://github.com/pmercado625/stock-analysis/blob/main/stock_performance_2017.png?raw=true)
![2018 stocks results](https://github.com/pmercado625/stock-analysis/blob/main/stock_performance_2018.png?raw=true)  


In comparing the success of this group of stocks, it's very apparent that most stocks ended up doing worse for themselves with the exceptions of ENPH and RUN which experienced fantastic growth from 2017 to 2018.   

![2017 Speed Test](https://github.com/pmercado625/stock-analysis/blob/main/VBA_Challenge_2017.png?raw=true)
![2018 Speed Test](https://github.com/pmercado625/stock-analysis/blob/main/old_code_2017.png?raw=true)  
![2017 Speed Test](https://github.com/pmercado625/stock-analysis/blob/main/VBA_Challenge_2018.png?raw=true)
![2018 Speed Test](https://github.com/pmercado625/stock-analysis/blob/main/old_code_2018.png?raw=true)  
Above is a side by side comparison of the new code to old code, respectively. As one can see, there is an entire magnitude of difference in time in favor of the refactored code to complete the same task.  
  
This difference in speed is likely due to the lack of the embedded for loops within the refactored code. By removing the need for the embedded for loops and simply going throug the entire data set once, it exponentially decreases the amount of actions needed to be taken by the program.  

Below is a code excerpt which shows the method used to iterate through each ticker
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        ' if the ticker above does not match the current ticker then...
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            'record the current row's price as the starting price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
                
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'record the current row's price as the ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            

            '3d Increase the tickerIndex.
            'increment the ticker to record for the next stock ticker
            tickerIndex = tickerIndex + 1  
        
 By having an index to reference an array, we can keep track of which ticker we're recording data for without having to iterate through the data set each time.





## Summary
