# Renewable Energy and VBA

## Overview of Project

### Purpose

The green wave is among us. Renewable energy has made significant strides with the advancing technology and the magnification of an environmentally conscious populace. Steve wants to capitalize on the bullish outlook of stocks related to green energy for the potential gains and his desire to join the movement towards a more sustainable future. With his vision of sustainability in mind, we will be analyzing a variety of green energy stocks for potential investment opportunities for Steve and his family. We will be using VBA in excel to automate our analysis so the end-user,i.e., Steve, can run complex analyses while mitigating errors, all with the press of a button. The future has never looked so bright and green! Perhaps there is a lucrative solar company that has what it takes to catch all that bright light. 

## Results

<img src = "resources/2017.png" width = 200> vs <img src = "resources/2018.png" width = 200>

### Stock Analysis

When comparing the stock analysis data of 2017 to 2018’s output, it becomes abundantly clear that 2018 was a tumultuous year for the stocks we analyzed. Only two companies had a positive return in 2018, where the returns for those two stocks were over 80 percent. Steve’s parents were interested in DQ’s stock, and while it had a phenomenal performance in 2017 with a 199.4% return, the following year amounted to a negative 62.6% return.  Alternatively, RUN had a minuscule growth in 2017 with only a 5% return but had the highest return in 2018 with an 84% return and one of the highest trading volumes. ENPH arguably performed the best with a 129.5% return in 2017 and an 81.9% return in 2018, along with the highest volumes of 2018. It seems as people are lining up to buy RUN and ENPH in 2018. Perhaps most of the stocks we analyzed were hyped up during the calendar year of 2017, which led to tremendous gains that year. Further analysis is required to determine the cause of the significant dips in 2018.

### Original Script Analysis

When comparing the processing times, the refactored script ran immensely quicker than the original script. The main difference between each script was the different implementation of arrays and nested loops. The original script only used one array to store the various stock tickers and used a nested for-loop. The nested loop's outer loop ran through each ticker and output the data for each ticker, while the inner loop used each ticker to produce the output data with conditional statements. The following code block displays the original script's for-loop initialization:

'''

    '4) Loop through tickers
    For I = 0 To 11
       ticker = tickers(I)
       totalVolume = 0
       
'''

After the for-loop is initialized, we nested the following loop within the loop above to produce the output data by having the conditional statements check if the current row's ticker matched the ticker(I) as the loop processed each individual row:

       '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For J = 2 To RowCount
       

### Refactored Script Analysis
The refactored script used four different arrays to store the stock tickers and their respective output data for the output sheet analysis. Stock tickers were accessed with a new index variable that would increase incrementally after each for-loop to access each ticker in succession. For example, after tickerIndex(0) accessed the AQ ticker for analysis, the tickerIndex() would increase by 1 to access the CSIQ ticker until the loop reaches its end with the VSLR ticker. The following script shows the tickerIndex() variable used in VBA:

'''

    'Created a ticker Index that incrementally accesses each stock ticker after they are looped over
        Dim tickerIndex As Integer
        tickerIndex = 0

The end of the loop increases the tickerIndex variable if the current ticker does not match the ticker of the row below with the following conditional statement nested in the for-loop:

'''

    'Check if the current row is the last row with the selected ticker.
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
            
            If Cells(I + 1, 1).Value <> tickers(tickerIndex) And Cells(I, 1).Value = tickers(tickerIndex) Then
        
                tickerEndingPrices(tickerIndex) = Cells(I, 6).Value

            'Increase the tickerIndex so next ticker can be analyzed and stored in output arrays.
                 tickerIndex = tickerIndex + 1
            
            End If
    
    Next I
    
Both scripts had the following array to access each ticker as needed, but the refactored script went a step further by having arrays made for the volume, starting price, and ending price:

'''

    Dim tickers(12) As String
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   
Refactored script's additional arrays used to store output data:

'''

        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

Refactored script’s method of producing output data via arrays:

'''

    For I = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + I, 1).Value = tickers(I)
        Cells(4 + I, 2).Value = tickerVolumes(I)
        Cells(4 + I, 3).Value = tickerEndingPrices(I) / tickerStartingPrices(I) - 1
        
    Next I
    
When comparing the processing times, the refactored script ran immensely quicker than the original script, even though the refactored script seems marginally longer than the original. The images below display the difference in execution times between the original and refactored script, where the message boxes on the left belong to the original script. The message boxes on the right belong to the refactored script. 

<img src = "resources/Original_2017.png" width = 200> vs <img src = "resources/VBA_Challenge_2017.png" width = 200>

<img src = "resources/Original_2018.png" width = 200> vs <img src = "resources/VBA_Challenge_2018.png" width = 200>

## Summary

### Disadvantages of Refactoring

The most prominent disadvantage of refactoring code is the time-consuming process that can lead to frustration and discouragement. The daunting task seemed almost impossible, where one roadblock would lead to another. The refactoring process involves a profuse amount of trial and error, where one can break the code altogether. At one point, I compromised the information in my "2017" sheet by running an incorrect code during the refactoring process. I had no choice but to start over.

### Advantages of Refactoring

The prominent advantage of refactoring code is improving the efficiency and readability of the code. The ultimate purpose of refactoring is doing more with less to save memory and processing power for other coding opportunities. The original script had coding blocks that were hard to follow, especially with the nested for-loop embedded deep within the "All Stocks Analysis" macro subroutine. Nested loops can become excessively chaotic, especially when one is running a variable through it one at a time. The reader of the code can become disconcerted as a result. Instead of running loops within loops until we get lost in confusion, we can make arrays of the data we want to analyze and run it through a few individual loops to improve the processing speed and the user's peace of mind. Additionally, when the original script ran, it took considerably longer than the refactored code. I also noticed that the excel sheet would flicker as if using a tremendous amount of processing power. After I refactored the script, the execution was seamless.  Imagine if the analyses were more extensive in scope with even more stocks to be analyzed. The original script would not be ideal, especially when we consider the practice of stock trading. To trade stocks proficiently, one must access and process the vast information available rapidly to ensure optimal procurement prices. In a high stake situation, we want the most optimal results in the shortest time frame. 

