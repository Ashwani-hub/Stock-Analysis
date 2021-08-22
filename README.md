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
![image](https://user-images.githubusercontent.com/82815722/130373314-67df2bc9-c581-43da-ac48-3c50c93c4cbf.png)
![image](https://user-images.githubusercontent.com/82815722/130373370-c6f49a18-a455-4049-9804-49838c5bf6d3.png)



  


## Code refactoring:

Original code was refactored by adding extra array, which allowed to loop through the rows and increased the tickerVolume and add the volume fo rthe current stock tickers. 

	For original code, only StartingPrice and endingPrice were initialized and the code was looped through all the values resulting in longer calculation time. 
	For refactored code, we used both inner and outer loops which aided with faster calculation time.
	
	![image](https://user-images.githubusercontent.com/82815722/130373346-40da0b49-1b09-4923-a7f3-e88ece414bbe.png)
	![image](https://user-images.githubusercontent.com/82815722/130373381-e4c5e76c-86fa-4ee5-ba5b-39506826c1b7.png)


