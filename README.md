# iSeries Query Tool

The iSeries Query Tool is a Python application that provides a graphical user interface for querying an IBM iSeries (AS/400) database. It allows users to connect to the database, execute SQL queries, visualize and export results, and more.

Start Up
![image](https://github.com/JAMl3/iSQT/assets/97791913/73d20de7-450a-4337-8865-c9dfdc061fa3)

Login
![image](https://github.com/JAMl3/iSQT/assets/97791913/bfe567d2-c44f-4b25-9934-689a85257e7c)

Select database query
![image](https://github.com/JAMl3/iSQT/assets/97791913/e19b2e3d-9523-4bb5-9ac6-b458a800f8b1)

View results/ Reorder columns/ Copy all to clipboard/ Export to excel/ Search
![image](https://github.com/JAMl3/iSQT/assets/97791913/76fc9f49-9921-425e-83b4-0aea7649aada)


## Features

- Connect to an IBM iSeries (AS/400) database using credentials.
- Choose a table and columns for the query from a graphical interface.
- Execute SQL queries on the selected table.
- View query results in a scrollable text box with horizontal and vertical scrolling.
- Search for specific text within the query results.
- Move selected columns' positions within the layout.
- Export query results to an Excel spreadsheet.
- Copy query results to the clipboard.

## Dependencies

- Python 3.x
- tkinter (usually included with Python)
- openpyxl
- pandas
- pyodbc
- PIL (Python Imaging Library)
- ttkthemes

## Installation

1. Install Python 3.x on your system.
2. Open a terminal or command prompt and navigate to the directory containing the script.
3. Install the required dependencies using the following command:

```bash
pip install openpyxl pandas pyodbc pillow ttkthemes
```

4. Replace `"company_Logo.png"` with the actual path to your company's logo image (if you have one).
5. Update the following parameters in the script according to your IBM iSeries setup:
    - Replace `YOUR_SERVER_IP` with the actual IP address of your iSeries server.
    - Replace `YOUR_LIBRARY_NAME` with the name of your library on the iSeries server.
6. Run the script using the following command:

```bash
python script_name.py
```

## Usage

1. Run the script using the provided installation steps.
2. Click the "Connect" button to enter your database credentials.
3. Update the following parameters in the script according to your IBM iSeries setup:
    - Replace `YOUR_SERVER_IP` with the actual IP address of your iSeries server.
    - Replace `YOUR_LIBRARY_NAME` with the name of your library on the iSeries server.
4. Select a table from the dropdown list.
5. Choose columns for your query from the list. You can also reorder columns using the "Move Up" and "Move Down" buttons.
6. Click the "Show Results" button to execute the query and display results in the text box.
7. Use the search bar and "Search" button to find specific text within the results.
8. Use the "Next" button to navigate through search results.
9. Use the "Export To Excel" button to save query results as an Excel file.
10. Use the "Copy to Clipboard" button to copy query results to the clipboard.



