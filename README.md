# SQL-Python
A simple script to run SQL commands in python and ouput results to Excel 

### Pre requisites:
Must have Python and SQL. Also must install psycopg2 and xlsxwriter modules. To do this run 
```
pip install psycopg2
pip install xlsxwriter
```

### Functions
There are 3 simple functions. 

#### get_queries
This function takes in a text file containing the queries. Queries can 
be over multiple lines however they two queries cannot be on the same line. To assign a label to a query, use SQL -- comment on the line before the query and this will become the query label. The function returns two lists, the first one containing the queries and the second contains the labels

#### execute_queries
This function will execute the queries. It takes a list of queries to run as the input and will return a list containing the results.

#### write_to_excel
This function will create the excel file. Takes in the results and labels for the queries and will create an excel sheet where the 1st column is the row number (i.e 1,2,3), the 2nd column is the label and the 3rd column is the query result.
