# VBA-challenge
This code in VBA analyzes generated stock market data from an Excel Workbook with three sheets, each with a year's worth of stock data from multiple stocks.

Code: The code in "Multiple_Year_Stock_Data_Krohn" is made up of four different subs, each with a different purpose:
  Sub Ticker: takes each individual ticker name and outputs it to Column I in each sheet.
    * Last row code can be found: https://www.excelcampus.com/vba/find-last-row-column-cell/
  Sub Yearly_change calculates the change from the opening price at the beginning of the year to the closing price at the end of the year. 
    * Note: this sub often takes two runs to populate both the numerical values and the conditional formatting.
    * The code for grabbing the Year Open and Year Close values as well as finding the total yearly change came from @TheodoreMoreland on github
  Sub Stock_T calculates the total volume of each stock for that year.
  Sub Greatest calculates and outputs the stock that showed the Greatest Percent Increase, Greatest Percent Decrease, and Greatest Total volume for that year, with the corresponding value in that category.
    * During a tutoring session, tutor Rebecca Leeds assisted with the code for finding the greatest percent increase and decrease

Though using 4 subs for one data workbook is more time-consuming than using one sub, the code is reliable and runs correctly. 
