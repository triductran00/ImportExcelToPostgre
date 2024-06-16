# Report Task 1: DB Import
## Name: Tran Duc Tri
## Problem:
The customer provides 100 Excel files, each containing one table. Tables with the same structure but different data, are sent to the company.
How to import these 100 files into PostgreSQL as quickly as possible while preserving the original data from the Excel files? 

## Technolog Used:
* **Python**: For scripting the application logic.
* **Tkinter**: For building the GUI.
* **psycopg2**: For interacting with PostgreSQL database.
* **pandas**: For data manipulation and analysis.
* **sqlalchemy**: For SQL toolkit and Object-Relational Mapping (ORM).
* **openpyxl**: For reading Excel files.
* **PostgreSQL**: As the backend database. 
  
## Implementation Step:
1. **Database Connection Details**: Users input their PostgreSQL database connection details including database name, user, password, host, and port.
   
2. **Save Connection Checkbox**: Users can choose to save the connection details for future use.
   
3. **Start Button**: Initiates the connection to the PostgreSQL database and opens the main application window.
   
4. **Main Application Window**:
* Import Data Button: Allows users to import Excel files into the PostgreSQL database.
* Delete Button: Opens a window to select and delete tables from the database. 
   
5. **Import Data**:
* Users select Excel files to import.
* Each file is read, and its contents are converted into a DataFrame.
* SQL statements for table creation are generated based on DataFrame structure.
* Tables are created in the database, and data is inserted into them.

6. **Delete Tables**:
* Users select tables they want to delete.
* Selected tables are deleted from the database.

7. **Logout**:
* Log out the user and delete any saved database connection details if the connection was not previously ticked.
* Returns the user to the Database Connection Details window.
  
## Flow Chart
```mermaid
graph TD;
    A[Start] --> B[Enter Database Connection Details];
    B --> C[Save Connection?];
    C -->|Yes| D[Save Connection Details];
    D --> E[Proceed];
    C -->|No| E[Proceed];
    E --> F[Connect to PostgreSQL];
    F --> G[Main Application Window];
    G --> H[Import Data];
    G --> I[Delete Tables];
    G --> V[Log Out]
    V --> B[Enter Database Connection Details];
    H --> J[Select Excel Files];
    J --> K[Read Excel Files];
    K --> L[Convert to DataFrame];
    L --> M[Generate SQL for Table Creation];
    M --> N[Create Tables in PostgreSQL];
    N --> O[Insert Data into Tables];
    O --> P[Show Success or Error Message];
    I --> Q[Select Tables to Delete];
    Q --> R[Delete Selected Tables];
    R --> P[Show Success or Error Message];
    P --> T[End];
    

  


