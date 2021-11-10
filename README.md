# stock-analysis
Performing analysis on Stock Data with VBA to analyze the performance of different stocks

## Overview of Project

### Purpose

This project is a continued dive into Excel, but this time using VBA routines to analyze the sucess of various stocks using 2 worksheets for 2 differnet years that contain the daily data of the stocks. We parse through each row to figure out the start price and end price, and use those figures to calcualte the yearly returns to see how the stock performed. 
## Results

The results of the stock analysis are very informative. The first thing the refactored code does is set up the basic outline of the results spreadsheet as well as setting up the ticker array that will house each individual ticker's information that we will use in the calculations. 

<img width="496" alt="Screen Shot 2021-11-09 at 3 42 40 PM" src="https://user-images.githubusercontent.com/92888170/141023580-500b4000-1663-420c-a31d-2bb9640ac272.png">


The next part of the code creates the ticker index and simply sets it to zero, as well as creating the arrays that will hold the values for the startingPrices, endingPrices, and Volumes of each individual ticker. The for loop runs through each index's Volume value and sets it to zero before we get into adding all the volumes for each stock.  

<img width="496" alt="Screen Shot 2021-11-09 at 3 43 20 PM" src="https://user-images.githubusercontent.com/92888170/141023941-67615507-6fe4-4fbd-b304-ceffdd4b034c.png">

The next part of the code is the most crucial. This is the part that, using a for loop, for each row we figure out if the current row is the first, or last instance of the current ticker index. If its the first apperacne of the ticker index, we set the price of the 6th column (Close Price) of the sheet of the current year to the array tickerStartingPrices at the current index. If its the last appearacne of the ticker index, we set the Close Price of the sheet to the array tickerEndingPrices at the current index. Also, if the index is the last appeacne, we add 1 to the ticker value to prepare the index for next pass in the for loop.

<img width="1094" alt="Screen Shot 2021-11-09 at 3 43 49 PM" src="https://user-images.githubusercontent.com/92888170/141024373-d835aa70-25d0-4529-916a-64a5fcbedb1d.png">

The last bit of refactored code goes through all of the arrays that we created with the previous for loop to produce the outcomes that we are looking for to measure the perfomance of the stocks. 

<img width="567" alt="Screen Shot 2021-11-09 at 3 44 26 PM" src="https://user-images.githubusercontent.com/92888170/141024740-33812f71-7ca0-433d-8134-37ccecc977e2.png">

Once this process is complete for both years, we can see how the stocks ended up perfoming. 

<img width="567" alt="Screen Shot 2021-11-09 at 4 00 44 PM" src="https://user-images.githubusercontent.com/92888170/141026480-8d24a111-3169-4ae6-8436-75a0e417588a.png">

<img width="567" alt="Screen Shot 2021-11-09 at 4 00 06 PM" src="https://user-images.githubusercontent.com/92888170/141026558-f88775e7-2121-425d-b14d-d5b65e573888.png">

Looking at the results, it becomes immediatley apparent that the 2 stocks that performed the best both years were ENPH and RUN. These were the only two stocks whose return improved year over year. These would be the only 2 stocks that I would recommend beacuse even though there were other stocks who had great performances one year, DQ and FLSR as an example, the volatililty of the stocks on a year by year basis is not something I would invest in. The ticker TERP was the only stock to decrease in value year over year, so I would consider shorting the stock if the trend continues. 

This code was a refactoring of the one that we did in the module for the class. The refactored code took 0.125 seconds for 2017 compared to 0.5976562 seconds for the non refactored code. For 2018, the code took 0.1289062 seconds for the refactored code, compared to 0.59375 for the original. In both cases, the refactored code is quicker. 

<img width="266" alt="Screen Shot 2021-11-09 at 3 31 41 PM" src="https://user-images.githubusercontent.com/92888170/141027449-153639bb-952a-4ab4-8d86-9db13887c1ca.png"> <img width="266" alt="Screen Shot 2021-11-09 at 3 31 16 PM" src="https://user-images.githubusercontent.com/92888170/141027458-7564ca2a-3f35-421c-9b75-de088f321a8b.png">


## Summary

There are some pros and cons to refactoring code. One of the main pros of refactoring code is that you already know how the code works and you are just applying it to a different setting. For me, one of the hardest parts of coding is trying to make the code work, so if there is already a template that can be worked with it makes the process a lot easier. Already knowing how the code works and being able to apply it to different settings can greatly increase the situations that can be analyzed with code. You can do alot of similar things by simply refactoring old code. Some cons of refactoring could be that the code you are refactoring just isnt that good. You could be using an inefficeint program to run something over and over agian when there is a different, better way. Another con is that that it could take alot of time to figure out how your originoal code could be used in the refactored form.

For me, in this project, the pros and cons were very apparent. For the pros, I already knew that my code worked from the module, I just had to implement it in the correct way. It was conformting to not have to start from scratch on a project because I just started learning VBA. On the cons side, I learned the hard way that it sometimes is diffuclt to see how the code can be used in its refactored form. It took me a long time to understand how the code could be manipuluated to become faster and how using indexing could be implemented in the code. But overall, the refactored version is better than the original. 

