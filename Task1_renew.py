import os
import psycopg2
import pandas as pd
from sqlalchemy import create_engine, inspect, text
import tkinter as tk
from tkinter import messagebox, filedialog
from openpyxl import load_workbook
from psycopg2 import sql

# Global variable to store database connection details
db_details = {}

# Function to create a database connection
def create_db_connection():
    return psycopg2.connect(
        database=db_details['database'],
        user=db_details['user'],
        password=db_details['password'],
        host=db_details['host'],
        port=db_details['port']
    )

# Function to test PostgreSQL connection
def test_postgres_connection():
    try:
        conn = create_db_connection()
        conn.close()
        return True
    except Exception as e:
        print("Error:", e)
        return False

# Function to create table SQL
def create_table_sql(data, table_name, schema='public'):
    columns = []
    primary_keys = []

    for _, row in data.iterrows():
        column_name = row.iloc[1]   # Column B (index 1)
        data_type = row.iloc[3]     # Column D (index 3)
        max_length = int(row.iloc[4]) if not pd.isna(row.iloc[4]) else None  # Max Length
        required = row.iloc[6] == 'Y'  # Column G (index 6)
        primary_key = not pd.isna(row.iloc[8])  # Column I (index 8)

        if data_type.lower() == 'varchar' and max_length:
            column_def = f'"{column_name}" {data_type}({max_length})'
        else:
            column_def = f'"{column_name}" {data_type}'

        if required:
            column_def += ' NOT NULL'

        columns.append(column_def)

        if primary_key:
            primary_keys.append(f'"{column_name}"')

    columns_sql = ',\n    '.join(columns)
    primary_keys_sql = ', '.join(primary_keys)

    if primary_keys:
        create_table_sql = f'''
        CREATE TABLE {schema}.{table_name} (
            {columns_sql},
            PRIMARY KEY ({primary_keys_sql})
        );
        '''
    else:
        create_table_sql = f'''
        CREATE TABLE {schema}.{table_name} (
            {columns_sql}
        );
        '''

    return create_table_sql

# Function to import Excel file into PostgreSQL
def import_excel_to_postgresql(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        table_name = df.iloc[2, 0]  
        header = df.iloc[7]        
        data = df.iloc[8:]          
        data.columns = header
        data = data.dropna(how='all')

        create_table_sql_statement = create_table_sql(data, table_name)
        print(f"Generated SQL: {create_table_sql_statement}")  

        conn = create_db_connection()
        cur = conn.cursor()

        # Check if table already exists
        cur.execute(f"SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_schema='public' AND table_name='{table_name.lower()}');")
        table_exists = cur.fetchone()[0]

        if table_exists:
            answer = messagebox.askyesno("Confirmation", f"Table '{table_name}' already exists. Do you want to overwrite it?")
            if answer:
                cur.execute(f"DROP TABLE IF EXISTS public.{table_name} CASCADE")
                conn.commit()
            else:
                cur.close()
                conn.close()
                return False

        cur.execute(create_table_sql_statement)
        conn.commit()
        cur.close()
        conn.close()
        print(f"Table '{table_name}' created successfully.")
        return True
    except Exception as e:
        print(f"Failed to import file {file_path}: {e}")
        return False

# Function to handle Import Data button click
def handle_import():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    success_files = [] 
    error_files = []  
    if file_paths:
        for file_path in file_paths:
            import_success = import_excel_to_postgresql(file_path)
            if import_success:
                success_files.append(file_path)  
            else:
                error_files.append((file_path, f"Failed to import file {file_path}"))
        
        if success_files:
            success_message = "Successfully imported the following files:\n"
            for file_path in success_files:
                success_message += f"- {file_path}\n"
            messagebox.showinfo("Success", success_message)
        
        if error_files:
            error_message = "Failed to import the following files:\n"
            for file_path, err_msg in error_files:
                error_message += f"{err_msg}\n"
            messagebox.showwarning("Warning", error_message)

# Function to delete selected tables
def handle_delete():
    def delete_table(table_listbox, delete_window):
        selected_tables = table_listbox.curselection()
        if not selected_tables:
            messagebox.showwarning("Warning", "Please select at least one table to delete.")
            return

        tables_to_delete = [table_listbox.get(index) for index in selected_tables]
        answer = messagebox.askyesno("Confirmation", f"Are you sure you want to delete the following tables?\n{', '.join(tables_to_delete)}")
        if answer:
            try:
                # Connect to the database
                engine = create_engine(f"postgresql://{db_details['user']}:{db_details['password']}@{db_details['host']}:{db_details['port']}/{db_details['database']}")
                conn = engine.connect()
                trans = conn.begin()

                for table in tables_to_delete:
                    print(f"Deleting table: {table}")
                    delete_sql = text(f"DROP TABLE IF EXISTS public.{table} CASCADE")
                    result = conn.execute(delete_sql)
                    print(f"Result: {result}")

                trans.commit()
                conn.close()

                for index in reversed(selected_tables):
                    table_listbox.delete(index)
                messagebox.showinfo("Success", "Selected tables have been deleted successfully.")
                delete_window.destroy()
            except Exception as e:
                print(f"An error occurred while deleting tables: {str(e)}")
                messagebox.showerror("Error", f"An error occurred while deleting tables: {str(e)}")

    delete_window = tk.Toplevel()
    delete_window.title("Select Tables to Delete")

    print("Retrieving table names...")
    engine = create_engine(f"postgresql://{db_details['user']}:{db_details['password']}@{db_details['host']}:{db_details['port']}/{db_details['database']}")
    inspector = inspect(engine)
    table_names = inspector.get_table_names()

    table_listbox = tk.Listbox(delete_window, selectmode=tk.MULTIPLE)
    for table in table_names:
        table_listbox.insert(tk.END, table)
    table_listbox.pack(padx=80, pady=40)

    delete_button = tk.Button(delete_window, text="Delete Tables", command=lambda: delete_table(table_listbox, delete_window))
    delete_button.pack(padx=20, pady=10)

    delete_window.mainloop()

# Function to save database connection details and start the main application
def save_and_start():
    global db_details
    if save_connection_var.get():  # Only save if the checkbox is checked
        db_details = {
            'database': db_name_entry.get(),
            'user': db_user_entry.get(),
            'password': db_password_entry.get(),
            'host': db_host_entry.get(),
            'port': db_port_entry.get()
        }
        messagebox.showinfo("Success", "Database connection details saved successfully.")
    if test_postgres_connection():
        print("Successfully connected to PostgreSQL database.")
        messagebox.showinfo("Success", "Successfully connected to PostgreSQL database.")
        root.withdraw()  
        main_app()
    else:
        print("Failed to connect to PostgreSQL database.")
        messagebox.showerror("Error", "Failed to connect to PostgreSQL database with the provided details.")

# Function to save database connection details
# Load saved database connection details, if available
def load_saved_details():
    global db_details
    if db_details:
        db_name_entry.insert(0, db_details.get('database', ''))
        db_user_entry.insert(0, db_details.get('user', ''))
        db_password_entry.insert(0, db_details.get('password', ''))
        db_host_entry.insert(0, db_details.get('host', ''))
        db_port_entry.insert(0, db_details.get('port', ''))

load_saved_details()

# Function to save database connection details
def save_db_details():
    global db_details
    if save_connection_var.get():
        db_details = {
            'database': db_name_entry.get(),
            'user': db_user_entry.get(),
            'password': db_password_entry.get(),
            'host': db_host_entry.get(),
            'port': db_port_entry.get()
        }
        messagebox.showinfo("Success", "Database connection details saved successfully.")
    else:
        db_details = {}

# Global variable to store the main window
main_window = None

# Function to logout and clear database connection details
def logout():
    global db_details
    if not save_connection_var.get():  # Clear the connection details only if "Save Connection" is not checked
        db_details = {}
        # Clear entry fields
        db_name_entry.delete(0, tk.END)
        db_user_entry.delete(0, tk.END)
        db_password_entry.delete(0, tk.END)
        db_host_entry.delete(0, tk.END)
        db_port_entry.delete(0, tk.END)
    if main_window:
        main_window.destroy()  # Destroy the main window if it exists
    messagebox.showinfo("Success", "Logged out successfully.")
    root.deiconify()  # Show the database connection window again

# Main application function
def main_app():
    global main_window
    main_window = tk.Toplevel(root)
    main_window.title("Database Management Tool")

    frame = tk.Frame(main_window)
    frame.pack(pady=100, padx=100)

    import_button = tk.Button(frame, text="Import Data", command=handle_import)
    import_button.pack(side=tk.LEFT, padx=20)

    delete_button = tk.Button(frame, text="Delete", command=handle_delete)
    delete_button.pack(side=tk.LEFT, padx=20)

    logout_button = tk.Button(main_window, text="Logout", command=logout)
    logout_button.pack(pady=20)

# Create Tkinter GUI for database connection details
root = tk.Tk()
root.title("Database Connection Details")
root.geometry("350x360")  # Adjusted size to accommodate the new logout button

tk.Label(root, text="Database Name (postgres, etc.)").grid(row=0, column=0, padx=10, pady=10)
db_name_entry = tk.Entry(root)
db_name_entry.grid(row=0, column=1, padx=20, pady=10)

tk.Label(root, text="User (postgres, etc.)").grid(row=1, column=0, padx=10, pady=10)
db_user_entry = tk.Entry(root)
db_user_entry.grid(row=1, column=1, padx=20, pady=10)

tk.Label(root, text="Password").grid(row=2, column=0, padx=10, pady=10)
db_password_entry = tk.Entry(root, show='*')
db_password_entry.grid(row=2, column=1, padx=20, pady=10)

tk.Label(root, text="Host (localhost, etc.)").grid(row=3, column=0, padx=10, pady=10)
db_host_entry = tk.Entry(root)
db_host_entry.grid(row=3, column=1, padx=20, pady=10)

tk.Label(root, text="Port (5432, etc.)").grid(row=4, column=0, padx=10, pady=10)
db_port_entry = tk.Entry(root)
db_port_entry.grid(row=4, column=1, padx=20, pady=10)

# Checkbox to save database connection details
save_connection_var = tk.BooleanVar()
save_connection_checkbox = tk.Checkbutton(root, text="Save Connection", variable=save_connection_var)
save_connection_checkbox.grid(row=5, columnspan=2, pady=10)

# Button to start the main application
start_button = tk.Button(root, text="Start", command=lambda: [save_db_details(), root.withdraw(), main_app()])
start_button.grid(row=6, columnspan=2, pady=10)

root.mainloop()




