import pandas as pd
import mysql.connector
from mysql.connector import Error
from tkinter import Tk, filedialog

def get_excel_file_path():
    root = Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.asksaveasfilename(
        title="Save Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )

    root.destroy()  # Close the Tkinter window

    return file_path

def export_data_to_excel():
    # Replace these variables with your MySQL root user connection details
    root_host = 'localhost'
    root_user = 'root'
    root_password = 'root'

    # Change the database name to 'Report'
    database_name = 'Report'

    # Build the SQL query
    sql_query = '''
        SELECT
            Company_Name,
            District_Name,
            ROUND(SUM(Ms_Cy_Sales), 2) AS Total_Ms_Cy_Sales,
            ROUND(SUM(Ms_Ly_Sales), 2) AS Total_Ms_Ly_Sales,
            ROUND(
                ((SUM(Ms_Cy_Sales) - SUM(Ms_Ly_Sales)) / NULLIF(SUM(Ms_Ly_Sales), 0)) * 100,
                2
            ) AS Growth_Rate
        FROM
            input_table
        GROUP BY
            Company_Name, District_Name
        ORDER BY
            Company_Name, District_Name
    '''

    try:
        # Connect to MySQL as root
        root_conn = mysql.connector.connect(host=root_host, user=root_user, password=root_password)
        if root_conn.is_connected():
            cursor = root_conn.cursor()

            # Use the database
            cursor.execute(f"USE {database_name}")

            # Execute the SQL query
            cursor.execute(sql_query)

            # Fetch the results into a DataFrame
            result_df = pd.DataFrame(cursor.fetchall(), columns=[desc[0] for desc in cursor.description])

            # Get the Excel file path from the user
            excel_file_path = get_excel_file_path()

            # Export the DataFrame to Excel
            result_df.to_excel(excel_file_path, index=False)

            print(f'Data exported to Excel: {excel_file_path}')

    except Error as e:
        print(f"Error: {e}")

    finally:
        # Close the cursor and connection
        if root_conn.is_connected():
            cursor.close()
            root_conn.close()
            print('MySQL connection closed')

if __name__ == "__main__":
    export_data_to_excel()
