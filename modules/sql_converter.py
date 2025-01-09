import re
from colorama import Fore
from openpyxl import Workbook

def generate_excel_from_sql(sql_content):
    create_table_pattern = re.compile(
        r'CREATE TABLE IF NOT EXISTS\s+`(?P<table>\w+)`\s*\((?P<columns>.*?)\)\s*ENGINE=.*?;',
        re.S | re.I
    )
    
    insert_into_pattern = re.compile(
        r'INSERT INTO\s+`(?P<table>\w+)`\s*\((?P<columns>.*?)\)\s*VALUES\s*(?P<values>.*?);',
        re.S | re.I
    )
    
    create_tables = create_table_pattern.findall(sql_content)
    if not create_tables:
        print(Fore.RED + '✘' + Fore.WHITE + " | No CREATE TABLE statements found in the SQL file.")
        return
    
    tables = {}
    
    for table in create_tables:
        table_name = table[0]
        columns_str = table[1]
        columns = extract_columns(columns_str)
        tables[table_name] = {'columns': columns, 'rows': []}
    
    insert_statements = insert_into_pattern.findall(sql_content)
    if not insert_statements:
        print(Fore.RED + '✘' + Fore.WHITE + " | No INSERT INTO statements found in the SQL file.")
        return
    
    for insert in insert_statements:
        table_name = insert[0]
        columns_str = insert[1]
        values_str = insert[2]
        columns = [col.strip(' `') for col in columns_str.split(',')]
        values = extract_values(values_str)
        for value in values:
            row = parse_value_row(value)
            if len(row) != len(columns):
                continue
            row_dict = dict(zip(columns, row))
            if table_name in tables:
                tables[table_name]['rows'].append(row_dict)
            else:
                tables[table_name] = {'columns': columns, 'rows': [row_dict]}
    
    available_tables = ", ".join(tables.keys())
    print(Fore.LIGHTBLUE_EX + "Available tables: " + Fore.WHITE + available_tables)
    selected_table = input(Fore.LIGHTBLUE_EX + "→" + Fore.WHITE + " | Select a table to export to Excel: ").strip()
    
    if selected_table not in tables:
        print(Fore.RED + '✘' + Fore.WHITE + " | The selected table does not exist.")
        return
    
    table = tables[selected_table]
    
    available_columns = ", ".join(table['columns'])
    print(Fore.LIGHTBLUE_EX + "Available columns: " + Fore.WHITE + available_columns)
    choice = input(Fore.LIGHTBLUE_EX + "→" + Fore.WHITE + " | Do you want to select specific columns to export? (yes/no): ").strip().lower()
    
    if choice == 'yes':
        selected_columns_input = input(Fore.LIGHTBLUE_EX + "→" + Fore.WHITE + " | Enter column names separated by commas: ")
        selected_columns = [col.strip() for col in selected_columns_input.split(",") if col.strip() in table['columns']]
        if not selected_columns:
            print(Fore.RED + '✘' + Fore.WHITE + " | No valid columns selected. Exiting...")
            return
    else:
        selected_columns = table['columns']
    
    new_column_names = {}
    rename_choice = input(Fore.LIGHTBLUE_EX + "→" + Fore.WHITE + " | Do you want to rename any columns? (yes/no): ").strip().lower()
    
    if rename_choice == 'yes':
        for col in selected_columns:
            new_name = input(Fore.LIGHTBLUE_EX + f"→ {Fore.WHITE} | Enter new name for column '{col}' (or press Enter to keep the same): ")
            new_column_names[col] = new_name if new_name else col
    else:
        new_column_names = {col: col for col in selected_columns}
    
    wb = Workbook()
    ws = wb.active
    ws.title = selected_table
    
    ws.append([new_column_names[col] for col in selected_columns])
    
    for row in table['rows']:
        ws.append([row.get(col, None) for col in selected_columns])
    
    excel_filename = selected_table + '.xlsx'
    try:
        wb.save(excel_filename)
        print(Fore.GREEN + '✔' + Fore.WHITE + f" | Excel file '{excel_filename}' has been created successfully!")
    except Exception as e:
        print(Fore.RED + '✘' + Fore.WHITE + f" | Error saving Excel file: {e}")

def extract_columns(columns_str):
    columns = []
    parts = split_sql_by_comma(columns_str)
    for part in parts:
        part = part.strip()
        if part.upper().startswith('PRIMARY KEY') or \
           part.upper().startswith('KEY') or \
           part.upper().startswith('INDEX') or \
           part.upper().startswith('CONSTRAINT'):
            continue
        match = re.match(r'`?(\w+)`?\s+', part)
        if match:
            columns.append(match.group(1))
    return columns

def extract_values(values_str):
    pattern = re.compile(r'\)\s*,\s*\(')
    rows = pattern.split(values_str.strip())
    cleaned_rows = [row.strip('() ') for row in rows]
    return cleaned_rows

def parse_value_row(value):
    values = []
    current = ''
    in_quotes = False
    escape = False
    for char in value:
        if escape:
            current += char
            escape = False
            continue
        if char == '\\':
            current += char
            escape = True
            continue
        if char == "'":
            in_quotes = not in_quotes
            current += char
            continue
        if char == ',' and not in_quotes:
            cleaned = clean_value(current)
            values.append(cleaned)
            current = ''
        else:
            current += char
    if current:
        cleaned = clean_value(current)
        values.append(cleaned)
    return values

def split_sql_by_comma(sql):
    parts = []
    current = ''
    paren_level = 0
    in_quotes = False
    escape = False
    for char in sql:
        if escape:
            current += char
            escape = False
            continue
        if char == '\\':
            current += char
            escape = True
            continue
        if char == "'":
            in_quotes = not in_quotes
            current += char
            continue
        if char == '(' and not in_quotes:
            paren_level += 1
        elif char == ')' and not in_quotes:
            paren_level -= 1
        if char == ',' and paren_level == 0 and not in_quotes:
            parts.append(current)
            current = ''
        else:
            current += char
    if current:
        parts.append(current)
    return parts

def clean_value(val):
    val = val.strip()
    if val.upper() == 'NULL':
        return None
    if val.startswith("'") and val.endswith("'"):
        return val[1:-1].replace("''", "'")
    return val