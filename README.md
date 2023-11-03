# Coffee-Sales

I found some data on Coffee Sales from https://github.com/mochen862/excel-project-coffee-sales which looked like a good dataset to clean the data; I then produced a visualisation page in Microsoft Excel. In the excel file, there were three sheets, 'orders', 'customers' and 'products'. In the 
'orders' sheet, there were 5 columns which had data populated, 'Order ID', 'Order Date', 'Customer ID', 'Product ID' and 'Quantity'; there were also eight other columns which needed to be populated. The 'customers' and 'products' contained data which was all populated with data values. I will be using the 'customers' and 'products' data to help me using the 'customers' and 'products' columns to populate the data values in the 'orders' sheet. 

The first column which I needed to populate was the 'customers' column which I populated using a xlookup function. The xlookup function I used is as follows:

=XLOOKUP(C2, customers!$A$1:$A$1001, customers!$B$1:$B$1001,,0) 

C2 is the Customer ID
customers!$A$1:$A$1001 is the Customer ID column in the customers sheet

This function looks up the 'Customer ID' in the  'Customer ID' column in the customers sheet. It then returns the corresponding value in the 'Email' column in the customers sheet. If the email for the corresponding Customer ID is not found then there is no value returned. Furthermore, the 0 at the end of the function indicates that the match must be exact.

Next, I wanted to populate the values in the ‘Email’ column thus I performed another xlookup function however I altered it as there were some missing ‘Email’ values in the ‘customers’ sheet. As a result, I accommodated for this by running my xlookup function as an if statement whereby the function would not populate the cell if the value was 0 and if there was a value, the xlookup function would run as usual:

=IF(XLOOKUP(C2,customers!$A$1:customers!$A$1001,customers!$C$1:$C$1001,,0)=0,"",(XLOOKUP(C2,customers!$A$1:customers!$A$1001,customers!$C$1:$C$1001,,0)))

C2 is the Customer ID
customers!$A$1:$A$1001 is the Customer ID column in the customers sheet
customers!$C$1:$C$1001 is the Email column in the customers sheet

This function states that if the xlookup function returns a value of 0, return a blank cell, if not, return the original xlookup value.
For the ‘Country’ column, I used a XLOOKUP function like with the ‘Customer Name’ column, however, I decided to use the index function for the remaining columns in which I needed to populate the data. As a result, I utilised the following index/match function:

=INDEX(products!$A$1:$G$49, MATCH(orders!$D2,products!$A$1:$A$49,0), MATCH(orders!I$1,products!$A$1:$G$1,0))

products!$A$1:$G$49 is the entire table in the products sheet
orders!$D2 product ID with the column reference absolute but the row reference is relative
products!$A$1:$A$49 is the product id column in the products sheet
orders!I$1 is the coffee type column with the row reference absolute but the column reference is relative
products!$A$1:$G$1 is the column titles in the products sheet

The index function looks at an array (in our case, the table in the products sheet) and returns a value or reference of the cell at the intersection of a particular row and column. In our case, the column is a match function which looks up the product ID in the products array and is an exact match. The column number is a match function which looks up the coffe type in the column titles of the products sheet and is an exact match. This allows me to easily populate the coffe type, roast type, size and unit prize columns.  

I populated all the remaining columns with this formula except for the ‘Sales’ column, which I worked out by multiplying the ‘Unit Price’ and ‘Quantity’ columns together.
I noticed that the ‘Coffee Type’ column is given in three-letter abbreviations therefore I wanted to create a column with the coffee type’s full name which I referred to as ‘Coffee Type Name’. The function I used to populate this column was:

=IF(I2="Rob", "Robusta", IF(I2="Exc", "Excelsa", IF(I2="Ara", "Arabica", IF(I2="Lib","Liberica",""))))

I2 is coffee type

I also noticed that the ‘Roast Type’ column was populated with one-letter abbreviations for the column thus I created a new column called ‘Roast Type Name’ which had a similar if statement to the one shown above.

The ‘Order Date’ column was displayed in the form of dd/mm/yyyy which could’ve been confusing as it was difficult to know whether it was an American format or British; therefore, I changed it to the format of dd-mmm-yyyy (e.g., 05-Sep-2019). In addition, the ‘Size’ column didn’t have 
any units assigned to it so I adjusted the entire to the format of  0.0 kg (e.g., 1.5 kg). I then added $ values to the ‘Unit Price’ and ‘Sales’ columns. 

I wanted to see which customers had a loyalty card therefore I made a column named ‘Loyalty Card’ and populated it with the following formula:

=XLOOKUP([@[Customer ID]],customers!$A$1:$A$1001,customers!$I$1:$I$1001,,0)

customers!$A$1:$A$1001 is the Customer ID column in the customers sheet
customers!$I$1:$I$1001 is the Loyalty Card column in the customers sheet


These are all the functions which I used to populate the fields. I then, made a dashboard in which users can slice by ‘Roast Type Name’, ‘Size’, ‘Loyalty Card’ and users can also filter which time period they want to focus on. A screenshot of the dashboard can be seen below:

![image](https://github.com/AdamH489/Coffee-Sales/assets/122322345/4a870bf8-259f-48a8-bc50-4ef160475fb1)

The line graph and bar charts were made with the use of 'pivot tables' and I enabled all the slicers to affect both of the graphs which I made. 

