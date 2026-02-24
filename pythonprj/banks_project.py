# Code for ETL operations on Country-GDP data

# Importing the required libraries
# Code for ETL operations on Country-GDP data

import pandas as pd
import requests
from bs4 import BeautifulSoup
from datetime import datetime
import sqlite3

def log_progress(message):
    ''' This function logs the mentioned message of a given stage of the
    code execution to a log file. Function returns nothing'''
    
    timestamp_format = '%Y-%h-%d-%H:%M:%S' # Year-Monthname-Day-Hour-Minute-Second
    now = datetime.now() # Get current time
    timestamp = now.strftime(timestamp_format) 
    
    # Open the file in append mode ('a') to ensure each entry is on a new line
    with open("./code_log.txt", "a") as f: 
        f.write(timestamp + ' : ' + message + '\n')

def extract(url, table_attribs):
    ''' This function aims to extract the required
    information from the website and save it to a data frame. The
    function returns the data frame for further processing. '''
    
    # Fetch the webpage content
    html_page = requests.get(url).text
    data = BeautifulSoup(html_page, 'html.parser')
    
    # Initialize an empty DataFrame with the required attributes
    df = pd.DataFrame(columns=table_attribs)
    
    # Find the tables; the 'By market capitalization' table is typically the first one
    tables = data.find_all('tbody')
    rows = tables[0].find_all('tr')
    
    for row in rows:
        col = row.find_all('td')
        if len(col) != 0:
            # Task: Identify required table data (Bank Name and Market Cap)
            # Link for the bank name is usually in the second <td>
            bank_name = col[1].find_all('a')[1].contents[0]
            # Market cap is in the third <td>
            # Task: Remove '\n' and typecast to float
            market_cap = float(col[2].contents[0].strip())
            
            # Create a dictionary for the current row
            dict = {
                "Name": bank_name,
                "MC_USD_Billion": market_cap # Task: Rename column to MC_USD_Billion
            }
            
            # Append the row to the DataFrame
            df_new_row = pd.DataFrame(dict, index=[0])
            df = pd.concat([df, df_new_row], ignore_index=True)
            
    return df


def transform(df, csv_path):
    ''' This function accesses the CSV file for exchange rate
    information, and adds three columns to the data frame, each
    containing the transformed version of Market Cap column to
    respective currencies'''
    
    # Read the exchange rate CSV file
    exchange_rate_df = pd.read_csv(csv_path)
    
    # Convert the DataFrame to a dictionary where 'Currency' is the key and 'Rate' is the value
    # Example: {'EUR': 0.93, 'GBP': 0.8, 'INR': 82.95}
    exchange_rate = exchange_rate_df.set_index('Currency').to_dict()['Rate']
    
    # Add new columns by multiplying the USD value with the exchange rate
    # Rounding to 2 decimal places as per project requirements
    df['MC_GBP_Billion'] = [round(x * exchange_rate['GBP'], 2) for x in df['MC_USD_Billion']]
    df['MC_EUR_Billion'] = [round(x * exchange_rate['EUR'], 2) for x in df['MC_USD_Billion']]
    df['MC_INR_Billion'] = [round(x * exchange_rate['INR'], 2) for x in df['MC_USD_Billion']]
    
    return df

def load_to_csv(df, output_path):
    ''' This function saves the final data frame as a CSV file in
    the provided path. Function returns nothing.'''
    
    # Use the pandas to_csv method to save the file
    # index=False is typically used to avoid saving the row numbers as a separate column
    df.to_csv(output_path, index=False)


def load_to_db(df, sql_connection, table_name):
    ''' This function saves the final data frame to a database
    table with the provided name. Function returns nothing.'''
    
    # Use to_sql to write the dataframe to the database
    # if_exists='replace' ensures that if the table exists, it is overwritten
    # index=False prevents the dataframe index from being saved as a column
    df.to_sql(table_name, sql_connection, if_exists='replace', index=False)


def run_query(query_statement, sql_connection):
    ''' This function runs the query on the database table and
    prints the output on the terminal. Function returns nothing. '''
    
    print(f"Query: {query_statement}")
    query_output = pd.read_sql(query_statement, sql_connection)
    print(query_output)
    print('\n')


# --- Execution ---

# Required entities
url = 'https://web.archive.org/web/20230908091635/https://en.wikipedia.org/wiki/List_of_largest_banks'
table_attribs = ["Name", "MC_USD_Billion"]

# Log preliminaries (Task 1)
log_progress('Preliminaries complete. Initiating ETL process')

# Call extract function (Task 2)
df = extract(url, table_attribs)

# Print the resulting data frame
print(df)

# Path to the exchange rate CSV provided
csv_path = './exchange_rate.csv'

# Log the start of transformation (from end of Task 2)
# log_progress('Data extraction complete. Initiating Transformation process')

# Execute Transformation
df = transform(df, csv_path)

# Print the transformed dataframe to verify (e.g., check the 5th bank's INR value)
print(df)
print(f"Market Cap of the 5th bank in INR: {df['MC_INR_Billion'][4]}")

# Log completion of transformation
log_progress('Data transformation complete. Initiating Loading process')

# Define the output path for the processed data
output_csv_path = './Largest_banks_data.csv'

# Execute Task 4: Load to CSV
load_to_csv(df, output_csv_path)

# Log completion of Task 4
log_progress('Data saved to CSV file')

# Define database and table names
db_name = 'Banks.db'
table_name = 'Largest_banks'

# 1. Log: SQL Connection initiated
log_progress('SQL Connection initiated.')

# 2. Create the connection object
sql_connection = sqlite3.connect(db_name)

# 3. Call the load_to_db function
load_to_db(df, sql_connection, table_name)

# 4. Log: Data loaded to Database as table
log_progress('Data loaded to Database as table. Running the query')

# 1. Print all contents of the table
query_1 = "SELECT * FROM Largest_banks"
run_query(query_1, sql_connection)

# 2. Print only the average market capitalization (MC_USD_Billion)
query_2 = "SELECT AVG(MC_USD_Billion) FROM Largest_banks"
run_query(query_2, sql_connection)

# 3. Print only the names of the top 5 banks
query_3 = "SELECT Name from Largest_banks LIMIT 5"
run_query(query_3, sql_connection)

# 4. Log: Process Complete
log_progress('Process Complete.')

# 5. Close the SQLite connection
sql_connection.close()

# 6. Log: Connection Closed (Optional, but good practice)
# log_progress('Server Connection closed.')


