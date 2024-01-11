import pandas as pd
import mysql.connector
from mysql.connector import Error
import re
from tkinter import Tk, filedialog

def get_excel_file_path():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )

    root.destroy()  # Close the Tkinter window

    return file_path

def import_data_to_mysql(excel_file_path):
    # Replace these variables with your MySQL root user connection details
    root_host = 'localhost'
    root_user = 'root'
    root_password = 'root'

    # Change the database name to 'Report'
    database_name = 'Report'

    # Default table name
    table_name = 'Input_Table'

    # Connect to MySQL as root to create the database
    try:
        root_conn = mysql.connector.connect(host=root_host, user=root_user, password=root_password)
        if root_conn.is_connected():
            print(f'Connected to MySQL server as root: {root_conn.server_host}:{root_conn.server_port}, User: {root_conn.user}')
            root_cursor = root_conn.cursor()

            # Create the database if it does not exist
            create_db_query = f"CREATE DATABASE IF NOT EXISTS {database_name};"
            root_cursor.execute(create_db_query)

    except Error as e:
        print(f"Error: {e}")

    finally:
        # Close the root cursor and connection
        if root_conn.is_connected():
            root_cursor.close()
            root_conn.close()
            print('Root MySQL connection closed')

    # Connect to the newly created database
    try:
        conn = mysql.connector.connect(host=root_host, user=root_user, password=root_password, database=database_name)
        if conn.is_connected():
            print(f'Connected to MySQL database: {conn.server_host}:{conn.server_port}, User: {conn.user}')
            cursor = conn.cursor()

            # Check if the table already exists
            check_table_query = f"SHOW TABLES LIKE '{table_name}';"
            cursor.execute(check_table_query)
            table_exists = cursor.fetchone()

            if not table_exists:
                # Read the Excel file into a pandas DataFrame
                df = pd.read_excel(excel_file_path, sheet_name=0, header=0)

                # Sanitize column names
                df.columns = [re.sub(r'\W+', '_', str(col)) for col in df.columns]

                # Create a table with the sanitized columns as the DataFrame
                create_table_query = f"CREATE TABLE {table_name} ({', '.join([f'{col} TEXT' if col in ['column1', 'column2'] else f'{col} VARCHAR(255)' for col in df.columns])});"
                cursor.execute(create_table_query)

                # Insert data into the MySQL table using prepared statements
                for index, row in df.iterrows():
                    # Replace NaN values with None before inserting into the database
                    row = [None if pd.isna(value) else value for value in row]

                    columns = ', '.join(df.columns)
                    placeholders = ', '.join(['%s'] * len(df.columns))
                    insert_query = f"INSERT INTO {table_name} ({columns}) VALUES ({placeholders})"
                    cursor.execute(insert_query, tuple(row))

                # Commit the changes
                conn.commit()
                print(f'Data imported to MySQL table: {table_name} in database: {database_name}.')
            else:
                print(f'Table {table_name} already exists. Skipping data import.')

    except Error as e:
        print(f"Error: {e}")

    finally:
        # Close the cursor and connection
        if conn.is_connected():
            cursor.close()
            conn.close()
            print('MySQL connection closed')

if __name__ == "__main__":
    excel_file_path = get_excel_file_path()

    if excel_file_path:
        import_data_to_mysql(excel_file_path)
    else:
        print("No file selected. Exiting.")
