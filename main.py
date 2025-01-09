from colorama import init, Fore
from modules.filehandler import load_file
from modules.sql_converter import generate_excel_from_sql

init(autoreset=True)

def main():
    file_path = input(Fore.LIGHTBLUE_EX + "→" + Fore.WHITE + " | Enter the path to your SQL file: ").strip()
    sql_content = load_file(file_path, mode='r')
    if sql_content:
        generate_excel_from_sql(sql_content)
    else:
        print(Fore.RED + '✘' + Fore.WHITE + " | Failed to load file. Check the path and try again.")

if __name__ == "__main__":
    main()