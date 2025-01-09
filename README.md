# Instructions for Using the SQL to Excel Converter Program

## How the Program Works

This program allows users to upload SQL files, select tables and columns, and then generate an Excel file. Users can choose which columns to move, rename them, and specify the name of the table to create. The program creates a file "xlsx", which can be used in the Excel program.

## Excel to SQL

You can also use a similar program to convert your Excel files to SQL.
https://github.com/Dimiqhz/Excel2SQL

## Installation

1. **Ensure that you have Python installed** (version 3.9 or higher). You can download it from the [official Python website](https://www.python.org/downloads/).
2. **Clone the repository** with the program code to your computer:
    ```bash
    git clone https://github.com/Dimiqhz/SQL2Excel.git
    ```
3. **Navigate to the program directory**:
    ```bash
    cd SQL2Excel
    ```
4. **Install the required libraries**:
    ```bash
    pip install openpyxl colorama xlrd
    ```

## Running the Program

1. Open a terminal or command prompt.
2. Navigate to the directory where your Python script is located.
3. Run the program with the following command:
    ```bash
    python main.py
    ```

## Using the Program

1. **Loading the SQL File**:
    - The program will prompt you to enter the path to the SQL file (`.sql`). Ensure that the file exists at the specified path.

2. **Selecting a Table**:
    - After loading the file, the program will display a list of available tables extracted from the SQL file. Select one by entering its name.

3. **Selecting Columns**:
    - The program will ask if you want to select specific columns for export. Answer `yes` or `no`.
    - If `yes`, enter the names of the columns separated by commas. Only the selected columns will be considered for renaming and data export.

4. **Renaming Columns**:
    - The program will ask if you want to rename the selected columns. If yes, you can rename each selected column individually. Enter new names for each column or leave it blank to keep the original names.

5. **Generating Excel**:
    - After completing all prompts, the program will create an `.xlsx` file in the current directory containing the exported data from the selected table with the specified columns.

## Example

![Screenshot](screenshots/image.png)

## Notes

- Ensure that the SQL file contains `CREATE TABLE` and `INSERT INTO` statements, as these will be used to extract table names and column data.
- The program does not support complex SQL data types like dates or numeric formats. All values will be recorded as text in the Excel file.
- Errors during file loading or Excel creation will be displayed on the screen.

## Conclusion

This program simplifies the process of converting SQL data into Excel spreadsheets. Customize your data exports as needed, and ensure your SQL files contain properly formatted `CREATE TABLE` and `INSERT INTO` statements for accurate data extraction. For further assistance or issues, feel free to contact the project maintainers.
