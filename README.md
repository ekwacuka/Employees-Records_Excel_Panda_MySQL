# Excel_Panda_MySQL
### This project imports an Excel spreadsheet into Python Pandas Dataframe. 
### Conducts some data verifications and then imports the data into a MySQL Database. 
### It then Separates the Spreadsheet programmatically into 2 separate worksheets. 

   #### Sheet 1: Authentication 
   #### Sheet 2: Employee Records

##### Disclaimer - The datasets are generated through random logic in VBA. 
##### These are not real human resource data and should not be used for any other purpose other than testing.

### Import Libraries
* import pandas as pd
* import pymysql
* from sqlalchemy import create_engine
* import mysql.connector as mysql

### create dataframe from excel file
* df = pd.read_excel (r'C:\Users\Emily Kimani\Desktop\Adv_Python\Pandas\100_Records_1.xlsx')

### Create a MySQL database using Python
* mycursor = mydb.cursor()
* mycursor.execute("CREATE DATABASE IF NOT EXISTS 100RECORDS")

### Establishing connection to the database by creating an engine that connects to the MySQL database.

* engine = create_engine("mysql+pymysql://root:emilykimani@localhost/100records")
    
### Write df into the table in MySQL database  

* df.to_sql('100records', engine, index=False)
* df = pd.read_sql_query("SELECT * FROM 100records", engine)

### Verify that 100Records database and 100Records table is in MySQL Workbench
![image](https://user-images.githubusercontent.com/77937714/113582603-8fae7280-95f6-11eb-986e-c9a94e8f8bfe.png)

## To Separate the Spreadsheet programmatically into 2 separate worksheets, I used the following steps

#### create authenticaction dataframe from the 2 columns 

* authentication = pd.DataFrame(df, columns= ['User Name', 'Password'])

#### create employees_records dataframe from the remaining columns

* employee_records = pd.DataFrame(df, columns= ['Emp ID', 'Name Prefix', 'First Name', 'Middle Initial', 'Last Name', 'Gender', 'E Mail', "Father's Name", "Mother's Name", "Mother's Maiden Name", 'Date of Birth', 'Time of Birth', 'Age in Yrs.', 'Weight in Kgs.', 'Date of Joining', 'Quarter of Joining', 'Half of Joining', 'Year of Joining', 'Month of Joining', 'Month Name of Joining', 'Short Month', 'Day of Joining', 'DOW of Joining', 'Short DOW', 'Age in Company (Years)', 'Salary', 'SSN', 'Phone No.', 'Place Name', 'County', 'City', 'State', 'Zip', 'Region'])

#### Create a Pandas Excel writer using XlsxWriter as the engine.
* writer = pd.ExcelWriter(r'C:\Users\Emily Kimani\Desktop\Adv_Python\Pandas\100_Records_combo.xlsx', engine='xlsxwriter')

#### Write each dataframe to a different worksheet.

* df.to_excel(writer, sheet_name='100_Records')
* authentication.to_excel(writer, sheet_name='authentication')
* employee_records.to_excel(writer, sheet_name='employee_records')

#### Close the Pandas Excel writer and output the Excel file.
* writer.save()

