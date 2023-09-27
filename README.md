# VBA-challenge
VBA Homework

For this challenge I had to create a script that loops through all the stocks for one year and outputs the following information:
  -  The ticker symbol
  -   Yearly change from the opening price at the beginning of a given year to the             closing price at the end of that year.
  -  The percentage change from the opening price at the beginning of a given year to the       closing price at the end of that year.
  - The total stock volume of the stock


I used the .Activate function to apply the code to all worksheets simultaniously, my tutor Limei Hu reccomended that to streamline the coding She also advised me to  set "Cells(i,1).Value" and "Cells(i+1,1).Value" as variables to make it easier to apply in the formulas and functions of the code

I added the Len(cell) function to the color condtionals to avoid the blank cells being filled which I found from a google search

I found the Max and Min functions quickly on google, but needed some help to find the 
Match function which I used to apply the appropriate ticker symbol. Saad Khan, from askBCS helped me to identify  the correct code to apply here. Saad also helped me realize I needed to re-lable my Dim values as Double, instead of Long as I had previously had them.

