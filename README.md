# VBA-challenge
VBA Homework

For this challenge I had to create a script that loops through all the stocks for one year and outputs the following information:
  -  The ticker symbol
  -   Yearly change from the opening price at the beginning of a given year to the             closing price at the end of that year.
  -  The percentage change from the opening price at the beginning of a given year to the       closing price at the end of that year.
  - The total stock volume of the stock


I used the .Activate function to apply the code to all worksheets simultaniously

I added the Len(cell) function to the color condtionals to avoid the blank cells being filled

I found the Max and Min functions quickly on google, but needed some help to find the 
Match function which I used to apply the appropriate ticker symbol
