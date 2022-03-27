Attribute VB_Name = "Module1"
'Create a script that loops through all the stocks for one year and outputs the following information:
'
'  * The ticker symbol.
'
'  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'
'  * The total stock volume of the stock.
'
'**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.
'
'The result should match the following image:
'
'![moderate_solution](Images/moderate_solution.png)
'
'## Bonus
'
'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:
'
'![hard_solution](Images/hard_solution.png)
'
'Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year) just by running the VBA script once.
'
'## Other Considerations
'
'* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3 to 5 minutes.
'
'* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with one click of a button.


Sub StockTickerReport():

'Set CurrentSheet to be able to handle the looping of each worksheet
'within the workbook.

Dim CurrentSheet As Worksheet

'Loop through all of the worksheets in the active workbook
'and create the headers for each column/row needed to develop the stock report.
'The entire process for each calculation and assignment will take place again when the next worksheet is selected

For Each CurrentSheet In Worksheets
    CurrentSheet.Range("I1").Value = "Ticker"
    CurrentSheet.Range("J1").Value = "Yearly Change"
    CurrentSheet.Range("K1").Value = "Percent Change"
    CurrentSheet.Range("L1").Value = "Total Stock Volume"
    CurrentSheet.Range("O2").Value = "Greatest % Increase"
    CurrentSheet.Range("O3").Value = "Greatest % Decrease"
    CurrentSheet.Range("O4").Value = "Greatest Total Volume"
    CurrentSheet.Range("P1").Value = "Ticker"
    CurrentSheet.Range("Q1").Value = "Value"
    
   'Setup variables to hold the values used for calculations
   'and cell designation
   
   Dim Ticker_ID As String
   Dim Total_Ticker_Volume As Double
   Dim Year_Opening_Price As Double
   Dim Year_Closing_Price As Double
   Dim Annual_Change As Double
   Dim Annual_Change_Percent As Double
   Dim Highest_Ticker As String
   Dim Lowest_Ticker As String
   Dim Highest_Volume As Double
   Dim Lowest_Volume As Long
   Dim High_Change_Percent As Double
   Dim Lowest_Change_Percent As Double
   Dim Highest_Volume_Ticker As String
   Dim Last_row As Long
   Dim Report_Row As Long
   
   'Define a variable for the Last_row. This will be used
   'in the loops to determine the rows to loop through for the entire raw dataset in column A
   
   Last_row = CurrentSheet.Cells(Rows.Count, 1).End(xlUp).Row
   
   'Define a variable for the report rows. This will be based on the number of unique
   'tickers. We will start of row 2 to 'skip' the header of the row
   Report_Row = 2
   
   'Setup the staring points for the calculation variables. Since these variables are outside of
   'the loop, they can act as trackers
   High_Change_Percent = 0
   Lowest_Change_Percent = 0
   Total_Ticker_Volume = 0
   'Set starting point for the opening price tracker to begin
   Year_Opening_Price = CurrentSheet.Cells(2, 3).Value
   
   'Loop from the beginning of the current worksheet, starting on the
   'second row. This loop will complete the bulk of the work, including the
   'assignment of the unique ticker labels for the report, along with the calculations
   'for the bonus section to the right of the initial report
   Dim i As Long
   For i = 2 To Last_row
        'Check to see if the current Ticker ID in column A matches as the loop goes
        'through the rows. If it does not match, then we know we have located the
        'section of the dataset needed to evaluate that specific stock
       If CurrentSheet.Cells(i + 1, 1).Value <> CurrentSheet.Cells(i, 1).Value Then
           Ticker_ID = CurrentSheet.Cells(i, 1).Value
           Year_Closing_Price = CurrentSheet.Cells(i, 6).Value
           Annual_Change = Year_Closing_Price - Year_Opening_Price
           Annual_Change_Percent = (Annual_Change / Year_Opening_Price) * 100
           Total_Ticker_Volume = Total_Ticker_Volume + CurrentSheet.Cells(i, 7).Value
           
            'Assign the values for the Ticker ID and the Annual Change to the report
           CurrentSheet.Range("I" & Report_Row).Value = Ticker_ID
           CurrentSheet.Range("J" & Report_Row).Value = Annual_Change
           
           'Create the conditional formatting steps if Yearly_Change > 0 cell fill should be green
           'if the Yearly_Change <=0 then the cell fill should be red
           If (Annual_Change > 0) Then
               CurrentSheet.Range("J" & Report_Row).Interior.ColorIndex = 4
           ElseIf (Annual_Change <= 0) Then
               CurrentSheet.Range("J" & Report_Row).Interior.ColorIndex = 3
           End If
           'Place the values for the Annual Change Percent and Total Ticker Volume
           CurrentSheet.Range("K" & Report_Row).Value = (CStr(Annual_Change_Percent) & "%")
           CurrentSheet.Range("L" & Report_Row).Value = Total_Ticker_Volume
           
           'Add 1 to the Report_Row in order to move to the next row of the stock report
           Report_Row = Report_Row + 1
           
           'Resetting the counters for the Annual_Change and the Year_Closing_Price to prepare
           'to move to the next Ticker_ID when the loop is complete
           
           Annual_Change = 0
           Year_Closing_Price = 0
           Year_Opening_Price = CurrentSheet.Cells(i + 1, 3).Value
           
           'This section of the program will work on the bonus section of the report
           'Here we calculate the highest and lowest changes in the tickers over the
           'course of the year. This calculation is inside the loop so it can change
           'as the observed values of the ticker calculations change
           
           If (Annual_Change_Percent > High_Change_Percent) Then
               High_Change_Percent = Annual_Change_Percent
               Highest_Ticker = Ticker_ID
           ElseIf (Annual_Change_Percent < Lowest_Change_Percent) Then
               Lowest_Change_Percent = Annual_Change_Percent
               Lowest_Ticker = Ticker_ID
           End If

           If (Total_Ticker_Volume > Highest_Volume) Then
               Highest_Volume = Total_Ticker_Volume
               Highest_Volume_Ticker = Ticker_ID
           End If
           'Similar to the reset we did before for the closing price and annual change, we
           'reset the annual change percent and total ticker volume to prepare for the next iteration of the loop.
           
           Annual_Change_Percent = 0
           Total_Ticker_Volume = 0

           Else
                Total_Ticker_Volume = Total_Ticker_Volume + CurrentSheet.Cells(i, 7).Value

        End If
        
   Next i
           'Assign the locations for the bonus report data
           CurrentSheet.Range("Q2").Value = (CStr(High_Change_Percent) & "%")
           CurrentSheet.Range("Q3").Value = (CStr(Lowest_Change_Percent) & "%")
           CurrentSheet.Range("P2").Value = Highest_Ticker
           CurrentSheet.Range("P3").Value = Lowest_Ticker
           CurrentSheet.Range("Q4").Value = Highest_Volume
           CurrentSheet.Range("P4").Value = Highest_Volume_Ticker
           
'Move to the next worksheet and repeat the reporting process
Next CurrentSheet
End Sub


'Reference
''***************************************************************************************/
'*    Title: Stock_Analysis_with_VBA
'*    Author: ibaloyan
'*    Date: 2018
'*    Code version: 1.0
'*    Availability: https://freesoft.dev/program/163047389
'*
'***************************************************************************************/
'(Version 2.0) [Source code]. http://www.graphicsdrawer.com

