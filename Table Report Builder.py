from typing import Any
import pandas as pd
import os
from tkinter import filedialog, messagebox
import tkinter as tk
import webbrowser
import tkinter.font as tkFont
from pandas import DataFrame
import multiprocessing
from multiprocessing.managers import DictProxy
import sys
import ctypes
from subprocess import run

# Turn on or off multiprocessing
USE_MULTICORE_PROCESSING = False if sum([0 if i.lower() not in sys.argv else 1 for i in ['-n', '-multioff', '--n', '--multioff', '/n', '/multioff', 'n', 'multioff']]) > 0 else True
USE_STR_CONVERSION = False if sum([0 if i.lower() not in sys.argv else 1 for i in ['-nc', '-noconvert', '--nc', '--noconvert', 'nc', 'noconvert']]) > 0 else True
USE_CASE_INSENSITIVE = True if sum([0 if i.lower() not in sys.argv else 1 for i in ['-c', '-case', '--c', '--case', 'c', 'case']]) > 0 else False

received_number: int = 0

# Enable text color for enabled systems
def enable_ansi_escape_sequences() -> bool:
    STD_OUTPUT_HANDLE = -11
    ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
    ENABLE_PROCESSED_OUTPUT = 0x0001
    kernel32 = ctypes.windll.kernel32
    # Get the handle to standard output
    handle = kernel32.GetStdHandle(STD_OUTPUT_HANDLE)
    if handle == 0 or handle == -1:
        raise ctypes.WinError()

    # Get current console mode
    mode = ctypes.c_uint()
    if not kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
        raise ctypes.WinError()

    # Modify mode to enable ANSI escape sequences
    new_mode = mode.value | ENABLE_VIRTUAL_TERMINAL_PROCESSING | ENABLE_PROCESSED_OUTPUT
    if not kernel32.SetConsoleMode(handle, new_mode):
        raise ctypes.WinError()

    return True


if os.name == "nt":
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        enable_ansi_escape_sequences()


# Print at a specific column in the terminal
def print_at_column(column: int, text: str) -> None:
    column = 1 if column < 1 else column
    sys.stdout.write(f"\033[{column}G{text}")
    sys.stdout.flush()

# DataFrame to Array function for running without multiprocessing
def collect_data_series(doc: DataFrame, num: int, line_limit: int, cols1) -> list[tuple[list[str], None, None, None]]: # type: ignore
    num_rows_1: int = len(doc) if len(doc) <= line_limit or line_limit == 0 else line_limit
    row_limit_1 = not (len(doc) <= line_limit or line_limit == 0)
    per_update_1: int = num_rows_1 // 100 if num_rows_1 // 100 != 0 else 1
    return [([doc.iloc[i][j] for j in cols1], print(f"Collecting data from table {["1 Raw Data ", "2 Raw Data ", "1 Reordered", "2 Reordered"][num-3]}{f" using row limitation {num_rows_1}" if row_limit_1 else ""}: \x1b[91m{int((i/num_rows_1)*100)}%\x1b[0m    Row {i} / {num_rows_1}{" "*15}", end="\r") if i%per_update_1 == 0 or i==0 else None) for i in range(num_rows_1)] # type: ignore

# Helper function for printing with multiprocessing
def assign_dict(dictionary: dict[int, Any], key: int, value: Any) -> None:
    dictionary[key] = value

# DataFrame to Array function for running with multiprocessing
def collect_data(doc: DataFrame, num: int, line_limit: int, cols1, collect_dict: dict[int, list[tuple[list[str], None, None, None]]] | DictProxy[Any, Any]) -> list[tuple[list[str], Any, None]]: # type: ignore
    num_rows_1: int = len(doc) if len(doc) <= line_limit or line_limit == 0 else line_limit
    per_update_1: int = num_rows_1 // 100 if num_rows_1 // 100 != 0 else 1
    num2 = num+4

    return [([doc.iloc[i][j] for j in cols1], assign_dict(collect_dict, num2, [([f"{["Table 1 Raw Data", "Table 2 Raw Data", "Table 1 Reordered", "Table 2 Reordered"][num-3]}: \x1b[91m{(str(int((i/num_rows_1)*100)))[-2:]}%\x1b[0m" if i%per_update_1 == 0 or i==0 else ""], None)]), print_at_column(1, "Collecting Data") if i%per_update_1 == 0 or i==0 else None, print_at_column(19+25*(num-3), f"{collect_dict[num2][0][0][0]}") if i%per_update_1 == 0 or i==0 else None) for i in range(num_rows_1)] # type: ignore

# Manager function for collect_data to use in Process objects
def collect_data_manager(doc: DataFrame, num: int, line_limit: int, cols1, collect_raw_dict: dict[int, list[tuple[list[str], None]]] | DictProxy[Any, Any]) -> None: # type: ignore
    collect_raw_dict[num] = collect_data(doc, num, line_limit, cols1, collect_raw_dict) # type: ignore
    print_at_column(19+25*(num-3), f"{["Table 1 Raw Data", "Table 2 Raw Data", "Table 1 Reordered", "Table 2 Reordered"][num-3]}: \x1b[92m100%\x1b[0m")

# Convert DataFrame array objects to a list of strings for row comparisons 
def rows_to_str(ls: list[tuple[list[str], None, None, None]]) -> list[str]:
    if USE_STR_CONVERSION:
        return [' '.join(list(map(conditional_convert, [str(row2).strip().replace('NULL', '0') for row2 in row[0]]))) for row in ls]
    return [' '.join([str(row2).strip().replace('NULL', '0') for row2 in row[0]]) for row in ls]

# Convert DataFrame array objects to a list of strings for row comparisons
# Converts to float or int first to remove trailing or leading zeros that can show as artifacts from different sources

# Read spreadsheet file and return a Pandas DataFrame object
def read_table(file: str, extension: str) -> DataFrame:
    print("Reading data from files                                                                    ", end="\r")
    if extension == "xlsx":
        try:
            doc: DataFrame = pd.read_excel(file, sheet_name=0)  # type: ignore
        except PermissionError:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("Access Denied", "If the document is open in any program, save and close it then click retry. Otherwise click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
        except:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("An Unknown Error Occurred", "An unknown error occurred while reading the file.\nAttempt to resolve the error and click retry, or click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()

    elif extension == "csv":
        try:
            doc: DataFrame = pd.read_csv(file)
        except PermissionError:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("Access Denied", "If the document is open in any program, save and close it then click retry. Otherwise click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
        except:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("An Unknown Error Occurred", "An unknown error occurred while reading the file.\nAttempt to resolve the error and click retry, or click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
    else:
        exit()
    return doc

def read_table_series(file: str, extension: str, message: str) -> DataFrame:
    print(message, end="\r")
    if extension == "xlsx":
        try:
            doc: DataFrame = pd.read_excel(file, sheet_name=0)  # type: ignore
        except PermissionError:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("Access Denied", "If the document is open in any program, save and close it then click retry. Otherwise click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
        except:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("An Unknown Error Occurred", "An unknown error occurred while reading the file.\nAttempt to resolve the error and click retry, or click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()

    elif extension == "csv":
        try:
            doc: DataFrame = pd.read_csv(file)
        except PermissionError:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("Access Denied", "If the document is open in any program, save and close it then click retry. Otherwise click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
        except:
            print("An error occurred while reading. Resolve it in the message box                             ", end="\r")
            ans = messagebox.askretrycancel("An Unknown Error Occurred", "An unknown error occurred while reading the file.\nAttempt to resolve the error and click retry, or click cancel to exit")
            if ans:
                doc = read_table(file, extension)
            else:
                exit()
    else:
        exit()
    return doc

# Manager function for read_table to use in Process objects for multiprocessing
def read_table_manager(file: str, extension: str, num: int, files_dict: dict[int, DataFrame] | DictProxy[Any, Any]) -> None:
    files_dict[num] = read_table(file, extension)

def row_comparison(rows1_str: list[str], rows2_str: list[str], message: str) -> tuple[int, int, list[int]]:
    print(message, end="\r")
    # Find number of rows from raw document 1 that exist in raw document 2
    raw_idx1: list[int] = []
    raw_yes1: int = 0
    raw_no1: int = 0
    for i in range(len(rows1_str)):
        if rows1_str[i] in rows2_str:
            raw_yes1 += 1
        else:
            raw_no1 += 1
            raw_idx1.append(i)

    return (raw_yes1, raw_no1, raw_idx1)

def row_comparision_multi(rows1_str: list[str], rows2_str: list[str], num: int, manager: dict[int, Any] | DictProxy[Any, Any]) -> None:
    # Find number of rows from raw document 1 that exist in raw document 2
    raw_idx1: list[int] = []
    raw_yes1: int = 0
    raw_no1: int = 0
    total: int = len(rows1_str)

    per_update_1: int = total // 100 if total // 100 != 0 else 1

    for i in range(total):
        _ = print_at_column(19+num*20, ["r1 vs r2: ", "r2 vs r1: ", "o1 vs o2: ", "o2 vs o1: "][num] + "\x1b[91m" + str(int((i/total)*100)) + "\x1b[0m") if i%per_update_1 == 0 else None
        if rows1_str[i] in rows2_str:
            raw_yes1 += 1
        else:
            raw_no1 += 1
            raw_idx1.append(i)
    print_at_column(19+num*20, ["r1 vs r2: ", "r2 vs r1: ", "o1 vs o2: ", "o2 vs o1: "][num] + "\x1b[92m100%\x1b[0m")
    
    manager[num] = (raw_yes1, raw_no1, raw_idx1)

def conditional_convert(item: str) -> str:
    if USE_CASE_INSENSITIVE:
        item = item.lower()
    if USE_STR_CONVERSION:
        if item.lower() == 'null':
            return "0.0"
        try:
            return str(float(item.lower().strip()))
        except:
            return item
    return item

def main():
    # Disables the input box for max rows. True to enable, False to disable
    USE_INPUT_FOR_MAX_ROW = True
    # If not using manual max row input, change max rows here. Set to 0 for the full table by default
    LIMIT_BYPASS = 100000

    # Print Header 
    run(['cls' if os.name == 'nt' else 'clear'], shell=True)
    print("Table Report Builder")
    print("+---------------------------------------------------------------------------------------------+", end="\n\n")
    print("Enter number of rows into the input window                                                     ", end="\r")

    # Determine if script has more than 1 core available or is multiprocessing is turned off with the flag
    # Sets manager dictionaries
    if multiprocessing.cpu_count() > 1 and USE_MULTICORE_PROCESSING:
        multicore: bool = True
        manager = multiprocessing.Manager()
        collection_dict: dict[int, list[tuple[list[str], None, None, None]]] | DictProxy[Any, Any] = manager.dict()
        collection_dict[7], collection_dict[8], collection_dict[9], collection_dict[10] = [[([""], None, None, None)] for _ in range(4)]
        file_dict: dict[int, DataFrame] | DictProxy[Any, Any] = manager.dict()
        comp_dict: dict[int, tuple[int, int, list[int]]] | DictProxy[Any, tuple[int, int, list[int]]] = manager.dict()
    else:
        multicore: bool = False
        collection_dict: dict[int, list[tuple[list[str], None, None, None]]] | DictProxy[Any, Any] = {}
        collection_dict[7], collection_dict[8], collection_dict[9], collection_dict[10] = [[([""], None, None, None)] for _ in range(4)]
        file_dict: dict[int, DataFrame] | DictProxy[Any, Any] = {}
        comp_dict: dict[int, (tuple[int, int, list[int]])] | DictProxy[Any, tuple[int, int, list[int]]] = {}

    # If the input box is enabled, use this section to gather user input
    if USE_INPUT_FOR_MAX_ROW:
        messagebox.showinfo("Maximum Rows", "Enter Maximum Rows (Tables will be truncated to limit if there are more rows than the limit) \nEnter 0 or leave blank for unlimited")
        # Handle submissions. Allows integers 0:infinity and empty strings
        def on_submit():
            global received_number
            value = entry_var.get()
            if value == "":
                received_number = 0
                root.quit()
                return
            try:
                received_number = int(float(value))
                root.quit()
            except ValueError:
                messagebox.showerror("Error", "Invalid number entered.")
        # Prevents illegal characters in input box
        def validate_number(val: str) -> bool:
            if val == "":
                return True  # Allow empty (user still typing)
            try:
                if int(val) >= 0:
                    return True
                else:
                    return False
            except ValueError:
                return False
            
        # Tkinter window setup
        root: tk.Tk = tk.Tk()
        root.wm_attributes('-topmost', 1) # type: ignore
        root.title("Enter Maximum Rows (Tables will be truncated to limit if there are more rows than the limit) \n Enter 0 or leave blank for unlimited")
        root.protocol("WM_DELETE_WINDOW", on_submit)
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        font_size = screen_width // 175
        root.geometry(f"{screen_width//4}x{screen_height//8}")
        entry_var = tk.StringVar()
        vcmd = (root.register(validate_number), "%P")
        tk.Label(root, text="Enter Maximum Rows\nTables will be truncated to limit if there are more rows than the limit\n Enter 0 or leave blank for unlimited", font=tkFont.Font(family="Arial", size=font_size)).pack(pady=5)
        entry = tk.Entry(root, textvariable=entry_var, validate="key", validatecommand=vcmd, font=tkFont.Font(family="Arial", size=font_size))
        entry.pack(pady=5)
        tk.Button(root, text="Submit", command=on_submit, font=tkFont.Font(family="Arial", size=font_size)).pack(pady=10)
        root.mainloop()
        
        root.destroy()
        line_limit = int(float(entry_var.get())) if entry_var.get() != "" else 0
    
    # If not using manual input, set line_limit to the set max value
    else:
        line_limit = LIMIT_BYPASS

    # Window information for file dialog to be always on top
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1) # type: ignore

    # Variable declaration
    doc1: DataFrame = pd.DataFrame({'col': ['val1']})
    doc2: DataFrame = pd.DataFrame({'col': ['val1']})

    # File dialog for file 1. Quit program if user selects cancel
    file1 = filedialog.askopenfilename(parent=root, filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], title="Select Table 1")
    extension1: str = file1.split(".")[-1].lower()
    exit() if len(extension1.strip()) == 0 else None
    

    # File dialog for file 2. Quit program if user selects cancel
    file2 = filedialog.askopenfilename(parent=root, filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], title="Select Table 2")
    extension2: str = file2.split(".")[-1].lower()
    exit() if len(extension2.strip()) == 0 else None


    root.destroy()

    # Ensure extensions are valid files
    if extension1 not in ["xlsx", "csv"] or extension2 not in ["xlsx", "csv"]:
        messagebox.showerror("Error", "Invalid file format. Please select an Excel or CSV file.")
        exit()

    # Reprint header with selected files
    run(['cls' if os.name == 'nt' else 'clear'], shell=True)
    print(f"Table Report Builder: \"{file1.split('/')[-1] if file1.count('/') > 0 else file1.split('\\')[-1]}\"  vs  \"{file2.split('/')[-1] if file2.count('/') > 0 else file2.split('\\')[-1]}\"")
    print("+---------------------------------------------------------------------------------------------+", end="\n\n")

    # Pathway for using multiprocessing
    if multicore:
        print("\x1b[32m*MULTIPROCESSING ENABLED\x1b[0m: Run Script with flag \'n\' or \'multioff\' to disable            ")
        if USE_STR_CONVERSION:
            print("\x1b[32m*AUTO NUMBER CONVERSION ENABLED\x1b[0m: Run script with flag \'noconvert\' or \'nc\' to disable            ")
        else:
            print("\x1b[31m*AUTO NUMBER CONVERSION DISABLED\x1b[0m: Run script without flags to enable                     ")
        if USE_CASE_INSENSITIVE:
            print("\x1b[32m*CASE INSENSITIVITY ENABLED\x1b[0m: Run script with flag \'c\' or \'case\' to disable             ")
        else:
            print("\x1b[31m*CASE INSENSITIVITY DISABLED\x1b[0m: Run script without flags to enable                         ")
        print()
        print("Reading data from files                                                                    ", end="\r")

        # Builds processes to read files
        jobs: list[multiprocessing.Process] = []
        jobs.append(multiprocessing.Process(target=read_table_manager, args=(file1, extension1, 1, file_dict)))
        jobs.append(multiprocessing.Process(target=read_table_manager, args=(file2, extension2, 2, file_dict)))
        jobs[0].start()
        jobs[1].start()
        jobs[0].join()
        jobs[1].join()
        doc1 = file_dict[1]
        doc2 = file_dict[2]
        jobs[0].close()
        jobs[1].close()

        # ======================================================

        # Get table column names
        print("Parsing Columns                                                                            ", end="\r")
        cols1 = doc1.columns
        cols2 = doc2.columns


        # Convert DataFrame column type sto string for accurate comparisons 
        print("Converting all columns to string type                                                      ", end='\r')
        for i in cols1:
            doc1[i] = doc1[i].astype('string')
        for i in cols2:
            doc2[i] = doc2[i].astype('string')


        # Replace all NaN, Null, None, or empty strings with "NULL" for accurate comparisons
        print('Converting NaN, Null, or None                                                              ', end='\r')
        doc1.fillna('NULL', inplace=True)
        doc2.fillna('NULL', inplace=True)

        # Number of rows in each DataFrame
        # num_rows_1: int = len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit
        # num_rows_2: int = len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit

        # Max rows for each DataFrame. Can be different if table sizes are different
        row_limit_1 = not (len(doc1) <= line_limit or line_limit == 0)
        row_limit_2 = not (len(doc2) <= line_limit or line_limit == 0)

        # ======================================================
        # Get the columns that are shared between the tables and sort alphabetically
        same_columns: list[str] = sorted(list(map(str, list(set(cols1) & set(cols2)))))

        # Set up processing variables for reordered comparison
        swap_idx1: list[int] = []
        swap_yes1: int = 0
        swap_no1: int = 0
        swap_idx2: list[int] = []
        swap_yes2: int = 0
        swap_no2: int = 0
        has_shared_columns = len(same_columns) > 0
        rows1_swap: list[tuple[list[str], None, None, None]] = []
        rows2_swap: list[tuple[list[str], None, None, None]] = []
        rows1_swap_str: list[str] = []
        rows2_swap_str: list[str] = []

        print("                                                                          ", end="\r")
        
        # Build processes for raw data collection to lists
        jobs: list[multiprocessing.Process] = [multiprocessing.Process(target=collect_data_manager, args=(doc1, 3, line_limit, cols1, collection_dict))] # type: ignore
        jobs.append(multiprocessing.Process(target=collect_data_manager, args=(doc2, 4, line_limit, cols2, collection_dict))) # type: ignore
        
        # Check if there are shared columns before attempting to process
        if has_shared_columns:
            # If there are shared columns, reindex the documents using the shared and sorted columns
            doc1_swap = doc1.reindex(columns=same_columns)
            doc2_swap = doc2.reindex(columns=same_columns)

            # Build processes for collecting reindexed documents into arrays
            jobs.append(multiprocessing.Process(target=collect_data_manager, args=(doc1_swap, 5, line_limit, same_columns, collection_dict))) # type: ignore
            jobs.append(multiprocessing.Process(target=collect_data_manager, args=(doc2_swap, 6, line_limit, same_columns, collection_dict))) # type: ignore

        # Start raw collection processes
        jobs[0].start()
        jobs[1].start()

        # Start reindexed collection processes
        if has_shared_columns:
            jobs[2].start()
            jobs[3].start()

        # Wait for raw collection processes to complete their processing
        jobs[0].join()
        jobs[1].join()

        # Wait for reindexed collection processes to complete their processing 
        if has_shared_columns:
            jobs[2].join()
            jobs[3].join()
        print()

        # Gather output data from raw collection
        rows1: list[tuple[list[str], None, None, None]] = collection_dict[3]
        rows2: list[tuple[list[str], None, None, None]] = collection_dict[4]


        print("                                                                                           ", end="\r")
        

        # Convert raw collection array of rows to array of strings. Not parallel as the processes are incredibly quick for large amount of data
        print("Table 1 row collection to strings                                                          ", end="\r")
        rows1_str = rows_to_str(rows1)
        print("Table 2 row collection to strings                                                          ", end="\r")
        rows2_str = rows_to_str(rows2)
        # ======================================================
        raw_idx1: list[int] = []
        raw_yes1 = 0
        raw_no1 = 0
        
        raw_idx2: list[int] = []
        raw_yes2 = 0
        raw_no2 = 0

        

        # ============================================================
        # Start Row Comparison Jobs
        print("Row Comparisons                                                                            ", end="\r")

        jobs: list[multiprocessing.Process] = []
        jobs.append(multiprocessing.Process(target=row_comparision_multi, args=[rows1_str, rows2_str, 0, comp_dict]))
        jobs.append(multiprocessing.Process(target=row_comparision_multi, args=[rows2_str, rows1_str, 1, comp_dict]))

        # Processing for shared columns if there are any 
        if has_shared_columns:
            # Gather output data from reindexed collection
            rows1_swap = collection_dict[5]
            rows2_swap = collection_dict[6]
            # Convert row arrays to strings (Not parallel as the processes are incredibly quick)
            rows1_swap_str = rows_to_str(rows1_swap)
            rows2_swap_str = rows_to_str(rows2_swap)
            
            jobs.append(multiprocessing.Process(target=row_comparision_multi, args=[rows1_swap_str, rows2_swap_str, 2, comp_dict]))
            jobs.append(multiprocessing.Process(target=row_comparision_multi, args=[rows2_swap_str, rows1_swap_str, 3, comp_dict]))

        jobs[0].start()
        jobs[1].start()

        if has_shared_columns:
            jobs[2].start()
            jobs[3].start()
            jobs[2].join()
            jobs[3].join()

        jobs[0].join()
        jobs[1].join()

        raw_yes1, raw_no1, raw_idx1 = comp_dict[0]
        raw_yes2, raw_no2, raw_idx2 = comp_dict[1]

        if has_shared_columns:
            swap_yes1, swap_no1, swap_idx1 = comp_dict[2]
            swap_yes2, swap_no2, swap_idx2 = comp_dict[3]
        print()
            

    # Pathway if multiprocessing is not used
    else:
        print("\x1b[31m*MULTIPROCESSING DISABLED OR UNAVAILABLE: Run script without flags to enable\x1b[0m               ")
        print("\x1b[31m**not available if system only has 1 processor\x1b[0m                                             \n")
        if USE_STR_CONVERSION:
            print("\x1b[32m*AUTO NUMBER CONVERSION ENABLED\x1b[0m: Run script with flag \'noconvert\' or \'nc\' to disable            \n")
        else:
            print("\x1b[31m*AUTO NUMBER CONVERSION DISABLED\x1b[0m: Run script without flags to enable                     \n")
        if USE_CASE_INSENSITIVE:
            print("\x1b[32m*CASE INSENSITIVITY ENABLED\x1b[0m: Run script with flag \'c\' or \'case\' to disable            \n")
        else:
            print("\x1b[31m*CASE INSENSITIVITY DISABLED\x1b[0m: Run script without flags to enable                     \n")
        
        # Read File 1 into a Pandas DataFrame object
        doc1 = read_table_series(file1, extension1, "Reading data from file 1                                                                   ")

        # Read File 2 into a Pandas DataFrame object
        doc2 = read_table_series(file2, extension2, "Reading data from file 2                                                                   ")

        # Get columns from each DataFrame
        print("Parsing Columns                                                                            ", end="\r")
        cols1 = doc1.columns
        cols2 = doc2.columns

        # Convert all columns to string types to ensure accurate comparisons
        print("Converting all columns to string type                                                      ", end='\r')
        for i in cols1:
            doc1[i] = doc1[i].astype('string')
        for i in cols2:
            doc2[i] = doc2[i].astype('string')


        # Convert all NaN, Null, None, or empty string values to NULL for accurate comparisons
        print('Converting NaN, Null, or None                                                              ', end='\r')
        doc1.fillna('NULL', inplace=True)
        doc2.fillna('NULL', inplace=True)


        # Collect data from raw table 1
        print("Collecting Data from file 1                                                                ", end="\r")
        rows1: list[tuple[list[str], None, None, None]] = collect_data_series(doc1, 1, line_limit, cols1)
        
        # Collect data from raw table 2
        print("Collecting Data from file 2                                                                ", end="\r")
        rows2: list[tuple[list[str], None, None, None]] = collect_data_series(doc2, 2, line_limit, cols2)

        # Convert table 1 row array to strings
        print("Table 1 row collection to strings                                                          ", end="\r")
        rows1_str = rows_to_str(rows1)
        
        # Convert table 2 row array to strings
        print("Table 2 row collection to strings                                                          ", end="\r")
        rows2_str = rows_to_str(rows2)


        # Row comparison
        # Find number of rows from raw document 1 that exist in raw document 2
        raw_idx1: list[int] = []
        raw_yes1 = 0
        raw_no1 = 0
        raw_yes1, raw_no1, raw_idx1 = row_comparison(rows1_str, rows2_str, "Raw Row Comparison - Compare Table 1 with Table 2                                          ")

        # Find number of rows from raw document 2 that exist in raw document 1
        raw_idx2: list[int] = []
        raw_yes2 = 0
        raw_no2 = 0
        raw_yes2, raw_no2, raw_idx2 = row_comparison(rows2_str, rows1_str, "Raw Row Comparison - Compare Table 2 with Table 1                                          ")

        # Get columns that exist in both tables and order alphabetically without duplicates
        same_columns: list[str] = sorted(list(set(cols1) & set(cols2)))

        # Setup variables for reordered/reindex processing
        swap_idx1: list[int] = []
        swap_yes1: int = 0
        swap_no1: int = 0
        swap_idx2: list[int] = []
        swap_yes2: int = 0
        swap_no2: int = 0
        has_shared_columns = len(same_columns) > 0
        rows1_swap: list[tuple[list[str], None, None, None]] = []
        rows2_swap: list[tuple[list[str], None, None, None]] = []
        rows1_swap_str: list[str] = []
        rows2_swap_str: list[str] = []


        print("Recheck with only shared columns                                                           ", end="\r")
        # Only run if there are shared columns
        if has_shared_columns:
            print("Reindexing with only shared columns                                                        ", end="\r")

            # Reindex documents to have only mutual columns in the same order
            doc1_swap = doc1.reindex(columns=same_columns)
            doc2_swap = doc2.reindex(columns=same_columns)

            # Number of rows in each document
            # num_rows_1: int = len(doc1_swap) if len(doc1_swap) <= line_limit or line_limit == 0 else line_limit
            # num_rows_2: int = len(doc2_swap) if len(doc2_swap) <= line_limit or line_limit == 0 else line_limit

            # Determine if the number of rows is being limited
            row_limit_1 = not (len(doc1_swap) <= line_limit or line_limit == 0)
            row_limit_2 = not (len(doc2_swap) <= line_limit or line_limit == 0)

            # Determine how often to print a percent completion update
            # per_update_1: int = num_rows_1//100 if num_rows_1//100 > 0 else 1
            # per_update_2: int = num_rows_2//100 if num_rows_2//100 > 0 else 1

            # Collect data from reindexed document 1
            print("Collecting Reordered Table 1                                                           ", end="\r")
            # rows1_swap = [([doc1_swap.iloc[i][j] for j in same_columns], print(f"Collecting reordered data from table 1{" using row limitation" if row_limit_1 else ""}: {int((i/num_rows_1)*100)}%    Row {i} / {num_rows_1}{" "*15}", end="\r") if i%per_update_1 == 0 else None, None, None) for i in range(num_rows_1)]
            rows1_swap = collect_data_series(doc1_swap, 2, row_limit_1, same_columns)
            
            # Collect data from reindexed document 2
            print("Collecting Reordered Table 2                                                           ", end="\r")
            # rows2_swap = [([doc2_swap.iloc[i][j] for j in same_columns], print(f"Collecting reordered data from table 2{" using row limitation" if row_limit_2 else ""}: {int((i/num_rows_2)*100)}%    Row {i} / {num_rows_2}{" "*15}", end="\r") if i%per_update_2 == 0 else None, None, None) for i in range(num_rows_2)]
            rows2_swap = collect_data_series(doc2_swap, 3, row_limit_2, same_columns)

            # Convert reindexed document row arrays to arrays of strings
            # rows1_swap_str = ["".join(str([row2.strip() for row2 in row[0]])) for row in rows1_swap]
            # rows2_swap_str = ["".join(str([row2.strip() for row2 in row[0]])) for row in rows2_swap]
            rows1_swap_str = rows_to_str(rows1_swap)
            rows2_swap_str = rows_to_str(rows2_swap)



            # print("Reordered Row Comparison - Compare Table 1 with Table 2                                ", end="\r")
            # Find number of rows from reordered document 1 that exist in reordered document 2
            # for i in range(len(rows1)):
            #     if rows1_swap_str[i] in rows2_swap_str:
            #         swap_yes1 += 1
            #     else:
            #         swap_no1 += 1
            #         swap_idx1.append(i)
            swap_yes1, swap_no1, swap_idx1 = row_comparison(rows1_swap_str, rows2_swap_str, "Reordered Row Comparison - Compare Table 1 with Table 2                                ")

            print("Reordered Row Comparison - Compare Table 2 with Table 1                                ", end="\r")
            # Find number of rows from reordered document 2 that exist in reordered document 1
            # for i in range(len(rows2)):
            #     if rows2_swap_str[i] in rows1_swap_str:
            #         swap_yes2 += 1
            #     else:
            #         swap_no2 += 1
            #         swap_idx2.append(i)
            swap_yes2, swap_no2, swap_idx2 = row_comparison(rows2_swap_str, rows1_swap_str, "Reordered Row Comparison - Compare Table 1 with Table 2                                ")
                    
        else:
            print("No shared columns found                                                                ", end="\r")


    # Dictionary used for building discrepancies table
    discrepancies: dict[str, list[str]] = {}
    

    # Compares individual cell values from the reindexed documents. In reindexing, rows remain in the same order
    # but non-mutual columns are removed, and shared columns are ordered
    # Shows a discrepancy for a cell even if it's row exists in the other table if the row is in a different position
    for i in range(len(same_columns)):
        # Set the value of column i in same_columns to a blank array which will represent the entirety of the column
        discrepancies[same_columns[i]] = []

        # Force height of the final discrepancy table to match the height of the shorter array
        # If table 1 is shorter, use the height of table 1 for discrepancy table
        if len(rows1_swap) < len(rows2_swap):
            for j in range(len(rows1_swap)):
                # Compare the cells at the given location [j][i] in the two tables
                if conditional_convert(str(rows1_swap[j][0][i]).strip()) != conditional_convert(str(rows2_swap[j][0][i]).strip()):
                    # If the cells are not equal, set row j in column same_columns[i] equal to 'cell 1 / cell 2'
                    discrepancies[same_columns[i]].append(str(rows1_swap[j][0][i]).strip() + " / " + str(rows2_swap[j][0][i]).strip())
                
                else:
                    # If the cells are equal, set row j in same_columns[i] equal to '.'
                    discrepancies[same_columns[i]].append(".")
        # If the heights are the same or table 2 is shorter, use the height of table 2
        else:
            for j in range(len(rows2_swap)):
                # Compare the cells at the given location [j][i] in the two tables
                if conditional_convert(str(rows1_swap[j][0][i]).strip()) != conditional_convert(str(rows2_swap[j][0][i]).strip()):
                    # If the cells are not equal, set row j in column same_columns[i] equal to 'cell 1 / cell 2'
                    discrepancies[same_columns[i]].append(str(rows1_swap[j][0][i]).strip() + " / " + str(rows2_swap[j][0][i]).strip())
                
                else:
                    # If the cells are equal, set row j in same_columns[i] equal to '.'
                    discrepancies[same_columns[i]].append(".")

    # Get plain file names for tables
    filename1 = os.path.basename(file1)
    filename2 = os.path.basename(file2)

    # Create a copy of each array of columns
    col1_cp = list(cols1)
    col2_cp = list(cols2)


    # Find non-mutual columns 
    for i in same_columns:
        del col1_cp[col1_cp.index(i)]
        del col2_cp[col2_cp.index(i)]

    print("Generating Report Document                                                                 ", end="\r")

    # Generate text for final document in HTML
    output = [
                # Metadata and header
                "<!DOCTYPE html>", 
                "<html lang=\"en\">", 
                "<head>", 
                "<meta charset=\"utf-8\">", 
                "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">", 
                # =============================================
                # CSS
                "<style>"+
                # p { line-height: 1rem; }
                "p\u007bline-height:1rem;\u007d"+

                # table {
                #   border-collapse: collapse;
                #   padding-right: 2em;
                # }
                "table\u007b"+
                "border-collapse:collapse;"+
                "padding-right:2em;\u007d"+

                # tr, td, th, table {
                #   border: 1px solid black;
                #   padding: 0.3rem 1rem 0.3rem 1rem;
                # }
                "tr, td, th, table\u007b"+
                "border: 1px solid black;"+
                "padding: 0.3rem 1rem 0.3rem 1rem;\u007d"+

                # .subsection { margin-left: 3em; }
                ".subsection\u007b margin-left: 3em;\u007d"+

                # .sectionend { margin-bottom: 3rem; }
                ".sectionend\u007bmargin-bottom:3rem;\u007d"+

                # .indent{ margin-left: 3em; }
                ".indent\u007b margin-left: 3em;\u007d"+

                # tr:nth-child(even) { background-color: #cccccc; }
                "tr:nth-child(even)\u007bbackground-color:#cccccc;\u007d"+

                # tr:nth-child(odd) { background-color: #eeeeee; }
                "tr:nth-child(odd)\u007bbackground-color:#eeeeee;\u007d"+

                # th { background-color: #66dddd }
                "th\u007bbackground-color:#66dddd\u007d"+
                "</style>", 
                # =============================================

                # Title
                "<title>Table Comparison</title>", "</head>", 
                
                # Body: style="font-size: 1.5rem; margin-right: 2em"
                "<body style=\"font-size: 1.5rem;margin-right:2em;\">",

                # Visual Title -- insert filenames of table 1 and 2 into filename1 and filename2
                # Defines position to jump to for "Back to Top" links
                f"<h1 id=\"top\"> Table Comparison - \"{filename1}\" and \"{filename2}\"</h1>",


                "<h2> File Information </h2>",
                "<div class=\"subsection\">",

                # Show names of files
                f"<p style=\"white-space:nowrap;\">File 1: {file1}</p>",
                f"<p style=\"white-space:nowrap;\">File 2: {file2}</p>",

                "</div>",

                # Quick Reference section (data summary)
                "<h2> Quick Reference </h2>",
                "<div class=\"subsection\">",

                # Show number of columns in each table
                f"<p> Raw Columns in Table 1: {len(cols1)}</p>",
                f"<p> Raw Columns in Table 2: {len(cols2)}</p>",

                # Show total number of rows in each table (rows processed. Shows line_limit if the document was truncated)
                f"<p>Total Rows in Table 1: {len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit}</p>",
                f"<p>Total Rows in Table 2: {len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit}</p>",

                # What percent of rows from table 1 exist in table 2 (Before reordering or using only shared columns)
                f"<p> Raw percent of rows in {filename1} that are in {filename2}: {round((raw_yes1 / (len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                f"<p> Raw percent of rows in {filename2} that are in {filename1}: {round((raw_yes2 / (len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",

                # Number of columns in common
                f"<p> Number of columns in common: {len(same_columns)}</p>",

                # What percent of rows from table 1 exist in table 2 (After reordering or using only shared columns)
                f"<p> Percent of rows in reordered {filename1} that are in reordered {filename2}: {round((swap_yes1 / (len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                f"<p> Percent of rows in reordered {filename2} that are in reordered {filename1}: {round((swap_yes2 / (len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",

                "</div>",

                # Links section 
                "<h2> Quick Links </h2>",

                # Jump to number of columns in table 1
                "<a href=\"#cc1\">Columns in Table 1</a> |",

                # Jump to number of columns in table 2
                "<a href=\"#cc2\">Columns in Table 2</a> |<br>",

                # Jump to Raw Row Comparision header
                "<a href=\"#rrc\">Raw Row Comparison</a> |",

                # Jump to Raw Row Comparison of table 1 into table 2
                "<a href=\"#rrc12\">Raw Row Comparison - Compare Table 1 with Table 2</a> |",

                # Jump to Raw Row Comparison of table 2 into table 1
                "<a href=\"#rrc21\">Raw Row Comparison - Compare Table 2 with Table 1</a> |<br>",

                # Jump to Reordered row comparison header
                "<a href=\"#rrcc\">Reordered Row Comparison</a> |",

                # Jump to reordered row comparison of table 1 into table 2
                "<a href=\"#rrcc12\">Reordered Row Comparison - Compare Table 1 with Table 2</a> |",


                # Jump to reordered row comparison of table 2 into table 1
                "<a href=\"#rrcc21\">Reordered Row Comparison - Compare Table 2 with Table 1</a> |<br>",

                # Jump to Column Value Discrepancy table
                "<a href=\"#cvd\">Column Value Discrepancies</a><br><br>",

                # Column Comparison section
                "<h2> Column Comparison </h2>",
                "<a href=\"#top\">Back to Top</a>",
                "<div class=\"subsection\">",
                
                # Show columns in table 1
                f"<p id=\"cc1\">Columns in Table 1: {len(cols1)}</p>",
                "<table>",
                "".join([f"<th>{i}</th>" for i in cols1]),
                "</table>",

                # Show columns in table 2
                f"<p id=\"cc2\">Columns in Table 2: {len(cols2)}</p>",
                "<table>",
                "".join([f"<th>{i}</th>" for i in cols2]),
                "</table>",

                "</div>",

                # Raw row comparison summary
                "<h2 id=\"rrc\"> Raw Row Comparison </h2>",
                "<div class=\"subsection\">",
                "<a href=\"#top\">Back to Top</a>",

                # Show number of rows in table 1
                f"<p>Total Rows in Table 1: {(len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit)}</p>",
                # Show number of rows in table 2
                f"<p class=\"sectionend\">Total Rows in Table 2: {(len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit)}</p>",


                # Raw Row Comparison: Comparing table 1 into table 2
                f"<h3 id=\"rrc12\"> Comparing {filename1} to {filename2}</h3>",

                "<a href=\"#top\">Back to Top</a>",

                # Number of rows in table 1 that exist in table 2
                f"<p class=\"indent\">Number of rows from {filename1} that appear in {filename2} -- {raw_yes1}</p>",
                # Number of rows in table 1 that do not exist in table 2
                f"<p class=\"indent\">Number of rows from {filename1} that do not appear in {filename2} -- {raw_no1}</p>",

                # Percent of rows in table 1 that are in table 2                               ( Number of rows from table 1 that appear in table 2  /  total number of rows in table 1 ) * 100, rounded to 2 decimals
                f"<p class=\"indent\"> Percent of rows in {filename1} that are in {filename2}: {round((raw_yes1 / (len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                
                # Show the rows from table 1 that do not appear in table 2
                f"<p class=\"indent\"> Rows from {filename1} that do not appear in {filename2}</p>",
                # Declare table object in DOM            Make "Row" column header
                "<table class=\"sectionend indent\"><tr><th>Row</th>",
                # Make table headers for each column in table 1
                "".join([f"<th>{i}</th>" for i in cols1]),
                "</tr>",
                # Data:                                   list comprehension to use an empty string to join a list of ["<td>a</td>", "<td>b</td>", ...]
                #              Row number in table                                         j represents columns index                        list of indexes of rows that don't exist
                "".join([f"<tr><td>{i+2}</td>" + "".join([f"<td>{rows1[i][0][j]}</td>" for j in range(len(rows1[i][0]))]) + "</tr>" for i in raw_idx1]),
                "</table>",


                # Raw Row Comparison: Comparing table 2 into table 1
                f"<h3 id=\"rrc21\"> Comparing {filename2} to {filename1}</h3>",
                "<a href=\"#top\">Back to Top</a>",

                # Number of rows in table 2 that exist in table 1
                f"<p  class=\"indent\">Number of rows from {filename2} that appear in {filename1} -- {raw_yes2}</p>",
                # Number of rows in table 2 that do not exist in table 1
                f"<p  class=\"indent\">Number of rows from {filename2} that do not appear in {filename1} -- {raw_no2}</p>",

                # Percent of rows in table 1 that are in table 2                               ( Number of rows from table 1 that appear in table 2  /  total number of rows in table 1 ) * 100, rounded to 2 decimals
                f"<p class=\"indent\"> Percent of rows in {filename2} that are in {filename1}: {round((raw_yes2 / (len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                
                # Show rows from table 2 that do not appear in table 1
                f"<p class=\"indent\"> Rows from {filename2} that do not appear in {filename1}</p>",
                "<table class=\"sectionend indent\"><tr><th>Row</th>",
                # Make table headers for each column in table 2
                "".join([f"<th>{i}</th>" for i in cols2]),
                "</tr>",
                # Data:                                   list comprehension to use an empty string to join a list of ["<td>a</td>", "<td>b</td>", ...]
                #              Row number in table                                         j represents columns index                        list of indexes of rows that don't exist
                "".join([f"<tr><td>{i+2}</td>" + "".join([f"<td>{rows2[i][0][j]}</td>" for j in range(len(rows2[i][0]))]) + "</tr>" for i in raw_idx2]),
                "</table></div>",

                # Reordered row table comparison header and Column Comparison summary 
                "<h2 id=\"rrcc\"> Reordered Row / Column Comparison </h2>",
                "<div class=\"subsection\">",
                "<a href=\"#top\">Back to Top</a>",
                "<h3> Summary </h3>",

                # Number of columns in table 1 that are not in table 2
                f"<p class=\"indent\"> Number of columns in Table 1 that do no appear in Table 2: {len(col1_cp)}</p>",
                f"<table class=\"sectionend indent\"><tr>",
                # Show columns in table 1 that are not in table 2
                "".join([f"<th>{i}</th>" for i in col1_cp]),
                "</tr></table>",

                # Number of columns in table 2 that are not in table 1
                f"<p class=\"indent\"> Number of columns in Table 2 that do no appear in Table 1: {len(col2_cp)}</p>",
                f"<table class=\"sectionend indent\"><tr>",
                # Show columns in table 2 that are not in table 1
                "".join([f"<th>{i}</th>" for i in col2_cp]),
                "</tr></table>",

                # Number of columns that appear in both tables
                f"<p class=\"indent\"> Number of columns that appear in both tables: {len(same_columns)}</p>",
                f"<p class=\"indent\"> Columns that appear in both tables:</p><table class=\"indent\"><tr>",
                # Show columns that appear in both tables
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr></table>",

                # Re show number of rows in each table
                f"<p class=\"indent\">Total Rows in Table 1: {len(doc1)}</p>",
                f"<p class=\"sectionend indent\">Total Rows in Table 2: {len(doc2)}</p>",

                # Reordered row comparison: table 1 into table 2
                f"<h3 id=\"rrcc12\"> Comparing reordered {filename1} to reordered {filename2}</h3>",
                "<a href=\"#top\">Back to Top</a>",
                # Number of rows from reordered table 1 that appear in reordered table 2
                f"<p class=\"indent\">Number of rows from reordered {filename1} that appear in reordered {filename2} -- {swap_yes1}</p>",
                # Number of rows from reordered table 1 that do not appear in reordered table 2
                f"<p class=\"indent\">Number of rows from reordered {filename1} that do not appear in reordered {filename2} -- {swap_no1}</p>",
                # Percent of rows in table 1 that are in table 2                                                       ( Number of rows from table 1 that appear in table 2  /  total number of rows in table 1 ) * 100, rounded to 2 decimals
                f"<p class=\"indent\"> Percent of rows in reordered {filename1} that are in reordered {filename2}: {round((swap_yes1 / (len(doc1) if len(doc1) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from reordered {filename1} that do not appear in reordered {filename2}</p>",
                "<table class=\"sectionend indent\"><tr><th>Row</th>",
                # Show rows from reordered table 2 that do not appear in reordered table 1
                # Make table headers for each column in table 2
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                # Data:                                   list comprehension to use an empty string to join a list of ["<td>a</td>", "<td>b</td>", ...]
                #              Row number in table                                         j represents columns index                        list of indexes of rows that don't exist
                "".join([f"<tr><td>{i+2}</td>" + "".join([f"<td>{rows1_swap[i][0][j]}</td>" for j in range(len(rows1_swap[i][0]))]) + "</tr>" for i in swap_idx1]),
                "</table>",

                # Reordered row comparison: table 2 into table 1
                f"<h3 id=\"rrcc21\">Comparing reordered {filename2} to reordered {filename1} </h3>",
                "<a href=\"#top\">Back to Top</a>",
                # Number of rows from reordered table 2 that appear in reordered table 1
                f"<p class=\"indent\" >Number of rows from reordered {filename2} that appear in reordered {filename1} -- {swap_yes2}</p>",
                # Number of rows from reordered table 2 that do not appear in reordered table 1
                f"<p class=\"indent\" >Number of rows from reordered {filename2} that do not appear in reordered {filename1} -- {swap_no2}</p>",
                # Percent of rows in table 2 that are in table 1                                                       ( Number of rows from table 1 that appear in table 2  /  total number of rows in table 1 ) * 100, rounded to 2 decimals
                f"<p class=\"indent\"> Percent of rows in reordered {filename2} that are in reordered {filename1}: {round((swap_yes2 / (len(doc2) if len(doc2) <= line_limit or line_limit == 0 else line_limit)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from reordered {filename2} that do not appear in reordered {filename1}</p>",
                "<table class=\"sectionend indent\"><tr><th>Row</th>",
                # Show rows from reordered table 2 that do not appear in reordered table 1
                # Make table headers for each column in table 2
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                # Data:                                   list comprehension to use an empty string to join a list of ["<td>a</td>", "<td>b</td>", ...]
                #              Row number in table                                         j represents columns index                        list of indexes of rows that don't exist
                "".join([f"<tr><td>{i+2}</td>" + "".join([f"<td>{rows2_swap[i][0][j]}</td>" for j in range(len(rows2_swap[i][0]))]) + "</tr>" for i in swap_idx2]),
                "</table></div>",


                # Column value discrepancies table
                "<h2 id=\"cvd\"> Column Value Discrepancies using value 1 / value 2 (Has false positives if rows are not in the same order in source tables)</h2>",
                "<a href=\"#top\">Back to Top</a>",
                "<table class=\"indent\"><tr>",

                # Row column                Header for each column in same_columns
                "<th>Row</th>" + "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                #        List comprehension builds each row
                #                   row number                         discrepancy dictionary value at row i in column j                                                                         If the whole row is equal (every cell is a '.') exclude it
                "".join([f"<tr><td>{str(i+2)}</td>" + "".join([f"<td> {discrepancies[j][i]} </td>" for j in same_columns]) + "</tr>" if len("".join([discrepancies[j][i] for j in same_columns]).replace(".", "")) != 0 else "" for i in range(len(discrepancies[same_columns[0]]))]),
                "<table style=\"user-select:none;-webkit-user-select:none;-ms-user-select:none;margin-left: 6em;opacity: 0;\"><tr>",
                "".join([f"<th>{i}</th>" for i in cols1])+"<th>space</th>",
                "</tr></table>",
                "</body>", 
                "</html>"
            ]

    # Find first valid file name
    filename = f"TableComparison_{filename1}_{filename2}"
    count = 1
    while os.path.exists(filename + ".html"):
        filename = f"TableComparison_{filename1}_{filename2}_{count}"
        count += 1

    # Write HTML to document
    file = open(filename + ".html", "w", encoding="utf-8") 
    file.write("\n".join(output))
    print(f"Table Comparison Report saved as {filename}.html")
    print("HTML report has been generated.")
    file.close()

    # Open generated document in default browser
    try:
        webbrowser.open(os.path.abspath(filename+'.html'))
    except Exception as _:
        messagebox.showerror("Error Opening File", "An error occurred when attempting to display the report. To open the report manually, locate the file \"" + os.path.abspath(filename+'.html') + "\"")





# Run main function
if __name__ == "__main__":
    main()
