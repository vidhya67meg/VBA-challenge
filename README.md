# VBA-challenge
Yearly stock analysis

[Input]

The excel sheet consisted of data such as open value, close value and volume of different stocks for each day and there are three sheets from the year 2018 to 2020. 

[Solution]

First, for the year 2018 I found the unique ticker symbols using a for loop in the vba script. Then within the for loop I determined the open value for the first date and closing value of the last date for each unique ticker to determine yearly change value and percentage change. I formatted both these values so that the increased values are green and decreased values are red. I also determined the total volume of each stock. Using conditionals I determined the maximum and minimum value of percentage change and maximum of total volume with its corresponding ticker symbol. Finally I added a for loop so the script runs for all the three sheets in the workbook.
