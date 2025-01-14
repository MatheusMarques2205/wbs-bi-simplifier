import pyodbc
import pandas as pd
from sqlalchemy import create_engine

# Database configuration
DATABASE_PATH = r"C:\Users\matheus.sales\OneDrive - Goldwind International\CS50 Harvard\Python Projects\WBS Database\EDF_Jacobina_Finame_WTG_Phase A_B.mdb"  # Replace with your database path
DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
CONN_STR = f"DRIVER={{{DRIVER}}};DBQ={DATABASE_PATH};"

# Function to determine the "Task" column based on the logic
def determine_task(row):
    if row['EDT1'] == 1 and row['EDT2'] == 1:
        return "01. NTP"
    elif row['EDT1'] == 2 and row['EDT2'] == 1 and pd.notnull(row['EDT3']):
        return "02. Anchor Cages EX Works"
    elif row['EDT1'] == 3 and row['EDT2'] == 1 and pd.notnull(row['EDT3']) and row['EDT4'] == 3:
        return "03. Anchor Cages Delivery at POD"
    elif row['EDT1'] == 4 and row['EDT2'] == 1 and pd.notnull(row['EDT3']) and pd.notnull(row['EDT4']) and pd.notnull(row['EDT5']):
        return "04. Anchor Cages Delivery at Site"
    elif row['EDT1'] == 2 and row['EDT2'] == 2 and row['EDT3'] == 1 and pd.notnull(row['EDT4']):
        return "05. WTG Ex Works Local"
    elif row['EDT1'] == 1 and row['EDT2'] == 2 and pd.notnull(row['EDT3']):
        return "07. Foundation Ready"
    elif row['EDT1'] == 4 and row['EDT2'] == 2 and pd.notnull(row['EDT3']) and pd.notnull(row['EDT4']):
        return "08. WTG Delivery at Site"
    elif row['EDT1'] == 5 and pd.notnull(row['EDT2']) and row['EDT3'] == 2:
        return "09. Main Crane Mobilization"
    elif row['EDT1'] == 5 and pd.notnull(row['EDT2']) and row['EDT3'] == 4 and pd.notnull(row['EDT4']) and row['EDT5'] == 3:
        return "10. Pre Erection"
    elif row['EDT1'] == 5 and pd.notnull(row['EDT2']) and row['EDT3'] == 5 and pd.notnull(row['EDT4']) and row['EDT5'] == 4:
        return "11. Full Erection"
    elif row['EDT1'] == 5 and pd.notnull(row['EDT2']) and row['EDT3'] == 3:
        return "12. Main Crane Demobilization"
    elif row['EDT1'] == 5 and pd.notnull(row['EDT2']) and row['EDT3'] == 6 and pd.notnull(row['EDT4']) and row['EDT5'] == 4:
        return "13. Eletrical and Mechanical Completion"
    elif row['EDT1'] == 6 and row['EDT2'] == 1:
        return "14. SCADA Installation"
    elif row['EDT1'] == 6 and row['EDT2'] == 2 and pd.notnull(row['EDT3']) and pd.notnull(row['EDT4']):
        return "15. Pre Commissioning"
    elif row['EDT1'] == 6 and row['EDT2'] == 3:
        return "16. Substation Energized"
    elif row['EDT1'] == 6 and row['EDT2'] == 4 and pd.notnull(row['EDT3']) and pd.notnull(row['EDT4']):
        return "17. Commissioning"
    elif row['EDT1'] == 6 and row['EDT2'] == 5 and pd.notnull(row['EDT3']) and pd.notnull(row['EDT4']):
        return "18. Run Test"
    elif row['EDT1'] == 7 and row['EDT2'] == 1 and row['EDT3'] == 2:
        return "19. COD"
    else:
        return 0

def main():
    try:
        # Connect to the database
        conn = pyodbc.connect(CONN_STR)
        print("Successfully connected to the database.")
        
        # Load the data into a pandas DataFrame
        query = "SELECT * FROM MSP_EpmTask_OlapView"
        df = pd.read_sql_query(query, conn)

        # Print columns for debugging
        print("Available columns:", df.columns.tolist())

        # Split the EDT column and apply task logic
        edt_split = df['EDT'].str.split('.', expand=True)
        for i in range(edt_split.shape[1]):
            df[f'EDT{i+1}'] = pd.to_numeric(edt_split[i], errors='coerce')
        df['Task'] = df.apply(determine_task, axis=1)

        # Create cursor and define table name
        cursor = conn.cursor()
        table_name = "MSP_EpmTask_OlapView_WBS_Output"
        
        # Create new table
        create_table_sql = f"""
        CREATE TABLE {table_name} (
            [TASK_ID] TEXT(255),
            [Task] TEXT(255)
        )
        """
        
        # Drop table if exists and create new one
        try:
            cursor.execute(f"DROP TABLE {table_name}")
        except:
            pass
        
        cursor.execute(create_table_sql)
        conn.commit()

        # Insert data
        for index, row in df.iterrows():
            if row['Task'] != 0:
                insert_sql = f"INSERT INTO {table_name} ([TASK_ID], [Task]) VALUES (?, ?)"
                cursor.execute(insert_sql, (row['TaskID'], row['Task']))
        
        conn.commit()
        print("Data saved to database successfully.")

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        if 'conn' in locals():
            conn.close()
            print("Database connection closed.")

if __name__ == "__main__":
    main()
