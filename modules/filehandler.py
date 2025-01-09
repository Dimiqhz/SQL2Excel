import os
from colorama import Fore

def load_file(file_path, mode='r'):
    if not os.path.isfile(file_path):
        print(Fore.RED + '✘' + Fore.WHITE + f" | File '{file_path}' does not exist.")
        return None
    try:
        with open(file_path, mode, encoding='utf-8') as file:
            return file.read()
    except Exception as e:
        print(Fore.RED + '✘' + Fore.WHITE + f" | Error reading file: {e}")
        return None