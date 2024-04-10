

### import some required packages

import mysql.connector
import csv
import re
import pymysql
import os
import pandas as pd


### Set up the path, list all file names, specify CSV.

csv_file_head = "D:/taiwan_stock_DailyQuotes_20040211_20240322_cleandata"
all_files = os.listdir(csv_file_head)
csv_files = [file for file in all_files if file.endswith('.csv')]

### Connect to MySQL database.

connection = mysql.connector.connect(host='127.0.0.1',
                                    port='3306',
                                    user='root',
                                    password='jj8879576')
cursor = connection.cursor()


### Design a loop to insert data into the database.

for csv_file in csv_files:
    full_path = os.path.join(csv_file_head, csv_file)
    
    # Convert the CSV file into matrix form, first output to TXT, and remove NaN values caused by data cleaning.
    
    df = pd.read_csv(full_path) 
    del tuple
    x = []
    y = []
    for i in range(len(df)):
        x = tuple(df.iloc[i])
        y.append(x)
    
    file = open(os.path.join(csv_file_head, "trytry.txt"),'w')
    for tuple in y:
        file.write(str(tuple) + ',' + '\n')
        
    file.close()
    
    with open(os.path.join(csv_file_head, "trytry.txt"), 'r') as file:
        lines = file.readlines()
    
    target_line = "(nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan, nan)"
    target_index = -1
    
    for i, line in enumerate(lines):
        if target_line in line:
            target_index = i
            break
        
    if target_index != -1:
        cleaned_lines = lines[:target_index]
    else:
        cleaned_lines = lines
        
    with open(os.path.join(csv_file_head, "trytry.txt"), 'w') as file:
        file.writelines(cleaned_lines)
    
    file.close()
    
    with open(os.path.join(csv_file_head, "trytry.txt"), 'r') as file:
        txt_content = file.read()
    
    
    # There are still some parts showing "NaN" due to lack of data. Save it in text format.
    last_comma_index = txt_content.rfind(',')  
    if last_comma_index != -1:  
        txt_content = txt_content[:last_comma_index] 

    txt_content = txt_content.replace('nan', '"nan"')
    

    # Split the data again and place it into a table
    table_name = os.path.splitext(csv_file)[0]
    
    sql_create_table = f"""CREATE TABLE `test`.`{table_name}` (`Security_Code` text,`Volume_Traded` text,
    `Number_of_Trades` text,`Transaction_Amount` text,`Opening_Price` text,`Highest_Price` text,
    `Lowest_Price` text,`Closing_Price` text,`Change_PlusMinus` text,`Price_Change` text,
    `Final_Bid_Price` text,`Final_Bid_Volume` text,`Final_Ask_Price` text,`Final_Ask_Volume` text,
    `Price_to_Earnings_Ratio` text)"""
    
    cursor.execute(sql_create_table)
    
    sql_insert = f"""INSERT INTO `test`.`{table_name}` VALUES{txt_content}"""
    cursor.execute(sql_insert)
    connection.commit()
    

