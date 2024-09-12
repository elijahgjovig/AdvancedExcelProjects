# Advanced Excel Project
I downloaded several data as seen in the xlsx file which I have linked. Used multiple different advanced functions and Excel features which I will include below

If you want to look in-depth at the full project, feel free to download the excel file.

## If Function
Utilized SumIf functions to determine if each salesperson met the weekly and monthly Sales Goals. Under "Bonus Status" I used the "IFAND" functions to first determine if the total sales of the salesperson for the month are equal to the top cell "I2" (34,000), AND second to determine if the cell range minimum is greater than the weekly goal of 8000 in Sales. Then the IF statement if true returns the value "BONUS" and if fales returns "NO BONUS" the formula I used looks like: 
=IF(AND($H$5="Yes",MIN(B5:E5)>=8000),"BONUS","NO BONUS")					

Lastly I used the COUNTIF function in order to count the number of people who received a bonus					

## SUMIF Function
On this sheet, I utilized the SUMIF Function to determine 2 things for 2 variables. First I wanted to find the total number of units sold for store # 3000. the function looks like =SUMIF(B2:B272,G3,E2:E272) 
then I did the same thing to determine the amount of $$$ in Sales for the same store. After that I did the same but for the specific stock-keeping unit specified in the G column. The equations remain the same, the only difference being the criteria section of the function.				

## VLOOKUPs
Utilized the ""VLOOKUP"" function to connect the employee ID on the ""VLOOKUP Function"" page to match the EMP ID accordingly with the employee's First Name, Last Name,  Department, and Pay Rate. In addition, I used the ""IFERROR"" function in combination with the VLOOKUPs to identify nonexistent Employee IDs. Function looks like:
=IFERROR(VLOOKUP($B18,'Master Emp List'!$A$1:$I$38,3,FALSE),""EMPID NOT FOUND"")
								
I noticed when I dragged the VLOOKUP Function over a column (for example First Name to Last Name), the function still indexes the same column number that I specified in the function syntax and returns the same value even though it's a different variable.  To combat this I combined both the INDEX and MATCH functions so that I only had to write one formula that serviced the entire table. In doing so the variables for every Employee ID were matched accordingly for this entire table.

=IFERROR(INDEX('Master Emp List'!$A$1:$I$38,MATCH($B23,'Master Emp List'!$A$1:$A$38,0),MATCH(C$2,'Master Emp List'!$A$1:$I$1,0)),""NO EMPLOYEE"")					
							
## HLOOKUPs
Used the HLOOKUP function, to look at the values in the table on the ""Master Inventory List"" spreadsheet. The lookup returns the number of units of inventory in each warehouse for the product code in cell B3 (XP200). Function looks like
=HLOOKUP($B$3,'Master Inventory List'!$A$2:$G$5,2,FALSE)			
					
Combined the MATCH and LEFT Functions with the HLOOKUP to match the ""row_index_num"" with the according Warehouse number in the HLOOKUP syntax, allowing the function to service the entire table seamlessly. Function looks like:
=HLOOKUP($B$3,'Master Inventory List'!$A$2:$G$5,MATCH(LEFT(A14,11),'Master Inventory List'!$A$2:$A$5,0),FALSE)		
					

## INDEX & MATCH Functions
Used the INDEX function to query the last name of an employee at a specific cell position in the "Master Inventory Spreadsheet." I used the Match function to tell me which specific row the EMP ID in the list on the left is located at on the in the "Master Inventory Spreadsheet		
						
Next, I used a combination of both the Index and Match Functions in combination with each other in order to index the table and match the EMP ID accordingly with that employee's department. Overcomes the limitations of both V & H Lookup functions. FUnction looks like:
=INDEX('INDEX MATCH Master Emp List'!$C$1:$C$38,MATCH(B4,'INDEX MATCH Master Emp List'!$D$1:$D$38,0))			
						
## LEFT RIGHT & MID Functions
Used the LEFT, RIGHT, and MID Functions to break up the SKU Number into 3 parts, Supplier ID, Part #, and Product Code. Functions Look Like:
=LEFT(A3,3)
=MID(A3,4,3)
=RIGHT(A3,2)
					
One of the SKU Numbers has a different format from the rest of the SKUs and has 10 characters instead of 8. to count the number of characters in the SKU Number I used the LEN Function. Function looks like
=LEN(A3)
					
## SEARCH Function
Used nesting of the SEARCH function into the LEFT Function to create a dynamic number.  The function looks like this:
=LEFT(A2,SEARCH("" "",A2))		
				
I noticed that the RIGHT Function reads left to right, so this function doesn't work to output the last name.
I used both the LEN and the SEARCH functions this time, nesting them with the RIGHT function. To create a dynamic number. The LEN function counts total characters, then I subtracted this by the SEARCH function which counts the characters until there is a space ("" ""). The equation looks like:
=RIGHT(A2,LEN(A2)-SEARCH("" "",A2))	
				
## CONCATENATE
All of the functions I've used so far are for extracting data. This time I combined text values using the CONCATENATE function. The equation looks like:
=CONCATENATE(C4,"" "",B4)
						
## PMT Function
Used PMT function to calculate the monthly payment based on the interest rate and the number of months. I found that the payment was too much for the allocated budget so I used the Goal Seek function to forecast the necessary mortgage amount to reach the target PMT
=-PMT(B3/12,B4,B2)
				
## Solver Add-in
The goal-seek function only allows you to optimize one variable. Because of this, I used the Solver Add-In to optimize the shipping cost based on several constraints.

1)  Each plant must produce at least 20 units a quarter
2)  Maximum capacity for Plant 1 is 92 units per quarter; Plant 2 = 45 units max; Plant 3 = 55 units max. 
3)  The plant production must equal the requirements of the warehouse

Based on the analysis and constraints, I found that the the minimum cost of shipping is $866.54
						
						
## Data Table
I Calculated the alternative interest rates in this data table. Still used the PMT function, and in addition, I used the Data table function in the What-If analysis			

## Scenario Manager
Used the scenario manager to create 3 different forecasts to predict the sales growth per quarter.				
				
## Intro to Macros & VBA
Stored code in a module to develop macros which resulted in seamlessly automating all future tables in Excel that I create.						
						
Tested the Macro with this set of data. The button above creates the Macro.					
					
