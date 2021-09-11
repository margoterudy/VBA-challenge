# VBA-challenge
VBA Assignment NUBootcamp
 In this folder use VBA scripting to analyze real stock market data.
 
 I tested my code on: https://nu.bootcampcontent.com/NU-Coding-Bootcamp/nu-chi-virt-data-pt-08-2021-u-c/-/blob/master/02-Homework/02-VBA-Scripting/Instructions/Resources/alphabetical_testing.xlsx
 
 Then performed my code on this following file. The excel file was too large to upload: https://nu.bootcampcontent.com/NU-Coding-Bootcamp/nu-chi-virt-data-pt-08-2021-u-c/-/blob/master/02-Homework/02-VBA-Scripting/Instructions/Resources/Multiple_year_stock_data.xlsx
 
 The instructions were as follows :
 Instructions

Create a script that will loop through all the stocks for one year and output the following information:

The ticker symbol.

Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

The total stock volume of the stock.

You should also have conditional formatting that will highlight positive change in green and negative change in red.

I started with naming the variables first. I then looped my worksheets so that all the tickers would be viewed by the code. I then created Columns with the correct titles. 
I was able to extract the tickers and measure the yearly change. It's important to note that the initial worksheet variables have to be set to zero for the correct numbers to populate.
I then set a loop to go through the ticker list and put the correct values in my cells.
I then ran If/Then to find the opening price and total stock volumes. 
I also set If/Then to recognize each time we hit a different ticker on the list. I then assigned the values in the correct columes
Color coding was set to Green if greater than Zero, Red if less than Zero, and Yellow if the value equated to Zero.
Note, it is very important to keep track of the If/Then commands by "End If" command.
I then set another If/Then code to extract the opening price and percent change per year. I then formated those cells. 
Finally, I found the Total Stock V in each worksheet (noting that each time we get to a new ticker it must be set to 0). 
Ended IF, then coded to cycle through a new worksheet. 

