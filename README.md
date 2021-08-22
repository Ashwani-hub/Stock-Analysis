# MechaCar_Statistical_Analysis

## Project Overview:

Using VBA to analyze the stocks for 2017 and 2018. 12 total tickers are available for analysis. Goal is to look at each ticker and calculate total volume and return percentage for each ticker by year. 

## Resources:
- Data Sources:
    - green_stocks.xlsx
    
- Software: 	Microsoft Office Home and Student 2019
		Excel (Version 2017: Build 14228.20250)
            
## Findings:

	Based on the calculation completed for 2017, it is observed that most of the tickers had a positive return and were safe investments.
	But after completing further analysis for year 2018, most of the stocks showed a huge fall.
	2 Best stocks that showed positive growth both in 2017 and 2018  were ENPH and RUN; with ENPH being best stock for both years.


  


## Code refactoring:

Original code was refactored by adding extra array, which allowed to loop through the rows and increased the tickerVolume and add the volume fo rthe current stock tickers. 

	For original code, only StartingPrice and endingPrice were initialized and the code was looped through all the values resulting in longer calculation time. 
	For refactored code, we used both inner and outer loops which aided with faster calculation time.
