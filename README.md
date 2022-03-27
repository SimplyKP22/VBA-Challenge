# VBA-Challenge
This repository showcases skills in VBA

## Overview:
	Before developing a report, an essential starting point is understanding the needs and expectations of your customer. Your customer could be internal or external, which could significantly change your approach. In this scenario, we assume that our customer is internal and therefore has a basic understanding of the data and will be able to understand the information presented in our report quickly. This assumption allows us to generate a direct report without many visualizations. We anticipate that the user will be comfortable moving around and manipulating the information. 

	In this example, we will be building a report using a dataset based on stock prices. We begin by establishing an excel (.xlsx) workbook containing the stock data we will analyze. Take a moment to understand the layout of the workbook and the data. Take note of the column headers and the data types that are located in each column. More specifically, some rows contain text data, and others have numbers. This will be an important consideration while developing your VBA script.
	
	We want each sheet to be populated with a general summary for each unique stock ticker within the first column for our report. We will also include a summary analysis of each ticker within the report. As a bonus step, we will construct a smaller report that provides information about the tickers with the greatest and lowest changes. This bonus report will be based on our initial report developed in each worksheet. NOTE: Once you begin working on your script, you will need to save the file as a macro-enabled workbook (.xlsm).
	
## Key Development Instructions:
Create a script that loops through all the stocks for one year and outputs the following information:

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

Use conditional formatting that will highlight positive change in green and negative change in red.

Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
