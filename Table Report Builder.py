import pandas as pd
import os
from tkinter import filedialog, messagebox
import tkinter as tk
import webbrowser
import tkinter.font as tkFont

recieved_number = 0

# If not using manual max row input, change max rows here
LIMIT_BYPASS = 1000

USE_INPUT_FOR_MAX_ROW = True

if os.name == "nt":
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(2)



def main():
    if USE_INPUT_FOR_MAX_ROW:
        messagebox.showinfo("Maximum Rows", "Enter Maximum Rows (Tables will be truncated to limit if there are more rows than the limit)")
        def on_submit():
            global recieved_number
            value = entry_var.get()
            if value == "":
                messagebox.showerror("Error", "Please enter a number.")
                return
            try:
                recieved_number = int(float(value))
                root.quit()
            except ValueError:
                messagebox.showerror("Error", "Invalid number entered.")
        def validate_number(val):
            if val == "":
                return False  # Allow empty (user still typing)
            try:
                val = int(val)
                if val > 0:
                    return True
                else:
                    return False
            except ValueError:
                return False
        root = tk.Tk()
        root.wm_attributes('-topmost', 1)
        root.title("Enter Maximum Rows (Tables will be truncated to limit if there are more rows than the limit)")
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        font_size = screen_width // 175
        root.geometry(f"{screen_width//8}x{screen_height//8}")
        entry_var = tk.StringVar()
        vcmd = (root.register(validate_number), "%P")
        tk.Label(root, text="Enter Maximum Rows:", font=tkFont.Font(family="Arial", size=font_size)).pack(pady=5)
        entry = tk.Entry(root, textvariable=entry_var, validate="key", validatecommand=vcmd, font=tkFont.Font(family="Arial", size=font_size))
        entry.pack(pady=5)
        tk.Button(root, text="Submit", command=on_submit, font=tkFont.Font(family="Arial", size=font_size)).pack(pady=10)
        root.mainloop()
        
        root.destroy()
        LINE_LIMIT = int(float(entry_var.get()))
    
    else:
        LINE_LIMIT = LIMIT_BYPASS

        
    root = tk.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)

    file1 = filedialog.askopenfilename(parent=root, filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], title="Select Table 1")
    file2 = filedialog.askopenfilename(parent=root, filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")], title="Select Table 2")

    root.destroy()


    extension1 = file1.split(".")[-1].lower()
    extension2 = file2.split(".")[-1].lower()
    if extension1 not in ["xlsx", "csv"] and extension2 not in ["xlsx", "csv"]:
        messagebox.showerror("Error", "Invalid file format. Please select an Excel or CSV file.")
        exit()

    print("Reading data from file 1", end="\r")
    if extension1 == "xlsx":
        doc1 = pd.read_excel(file1)
    elif extension1 == "csv":
        doc1 = pd.read_csv(file1)

    print("Reading data from file 2", end="\r")
    if extension2 == "xlsx":
        doc2 = pd.read_excel(file2)
    elif extension2 == "csv":
        doc2 = pd.read_csv(file2)

    print("Parsing Columns                                                                            ", end="\r")
    cols1 = doc1.columns
    cols2 = doc2.columns

    print('Converting NaN, Null, or None                                                              ', end='\r')
    doc1.fillna('NULL', inplace=True)
    doc2.fillna('NULL', inplace=True)

    for i in cols1:
        doc1[i] = doc1[i].astype('string')
    for i in cols2:
        doc2[i] = doc2[i].astype('string')

    print("Collecting Data from file 1                                                                ", end="\r")
    rows1 = [[doc1.iloc[i][j] for j in cols1] for i in (range(len(doc1)) if len(doc1) <= LINE_LIMIT else range(LINE_LIMIT))]
    print("Collecting Data from file 2                                                                ", end="\r")
    rows2 = [[doc2.iloc[i][j] for j in cols2] for i in (range(len(doc2)) if len(doc2) <= LINE_LIMIT else range(LINE_LIMIT))]

    rows1_str = [" ".join(str(row)) for row in rows1]
    rows2_str = [" ".join(str(row)) for row in rows2]

    # Row comparison

    print("Raw Row Comparison - Compare Table 1 with Table 2                                          ", end="\r")
    # 1 in 2
    raw_idx1 = []
    raw_yes1 = 0
    raw_no1 = 0
    for i in range(len(rows1_str)):
        if rows1_str[i] in rows2_str:
            raw_yes1 += 1
        else:
            raw_no1 += 1
            raw_idx1.append(i)

    print("Raw Row Comparison - Compare Table 2 with Table 1                                          ", end="\r")
    # 2 in 1
    raw_idx2 = []
    raw_yes2 = 0
    raw_no2 = 0
    for i in range(len(rows2_str)):
        if rows2_str[i] in rows1_str:
            raw_yes2 += 1
        else:
            raw_no2 += 1
            raw_idx2.append(i)

    same_columns = sorted(list(set(cols1) & set(cols2)))

    swap_idx1 = []
    swap_yes1 = 0
    swap_no1 = 0
    swap_idx2 = []
    swap_yes2 = 0
    swap_no2 = 0
    has_shared_columns = len(same_columns) > 0

    print("Recheck with only shared columns                                                           ", end="\r")
    if has_shared_columns:
        doc1_swap = doc1.reindex(columns=same_columns)
        doc2_swap = doc2.reindex(columns=same_columns)


        rows1_swap = [[doc1_swap.iloc[i][j] for j in same_columns] for i in (range(len(doc1_swap)) if len(doc1_swap) <= LINE_LIMIT else range(LINE_LIMIT))]
        rows2_swap = [[doc2_swap.iloc[i][j] for j in same_columns] for i in (range(len(doc2_swap)) if len(doc2_swap) <= LINE_LIMIT else range(LINE_LIMIT))]

        rows1_swap_str = ["".join(str(row)) for row in rows1_swap]
        rows2_swap_str = ["".join(str(row)) for row in rows2_swap]


        print("Reordered Row Comparison - Compare Table 1 with Table 2                                ", end="\r")
        # 1 in 2

        for i in range(len(rows1)):
            if rows1_swap_str[i] in rows2_swap_str:
                swap_yes1 += 1
            else:
                swap_no1 += 1
                swap_idx1.append(i)

        print("Reordered Row Comparison - Compare Table 2 with Table 1                                ", end="\r")
        # 2 in 1
        for i in range(len(rows2)):
            if rows2_swap_str[i] in rows1_swap_str:
                swap_yes2 += 1
            else:
                swap_no2 += 1
                swap_idx2.append(i)
    else:
        print("No shared columns found                                                                ", end="\r")

    discrepancies = {}

    for i in range(len(same_columns)):
        discrepancies[same_columns[i]] = []
        if len(rows1_swap) < len(rows2_swap):
            for j in range(len(rows1_swap)):
                if str(rows1_swap[j][i]) != str(rows2_swap[j][i]):
                    discrepancies[same_columns[i]].append(str(rows1_swap[j][i]) + " / " + str(rows2_swap[j][i]))
                else:
                    discrepancies[same_columns[i]].append(".")
        else:
            for j in range(len(rows2_swap)):
                if str(rows2_swap[j][i]) != str(rows1_swap[j][i]):
                    discrepancies[same_columns[i]].append(str(rows1_swap[j][i]) + " / " + str(rows2_swap[j][i]))
                else:
                    discrepancies[same_columns[i]].append(".")
    filename1 = os.path.basename(file1)
    filename2 = os.path.basename(file2)

    col1_cp = list(cols1)
    col2_cp = list(cols2)

    print(type(col1_cp))

    for i in same_columns:
        del col1_cp[col1_cp.index(i)]
        del col2_cp[col2_cp.index(i)]


    output = ["<!DOCTYPE html>", "<html lang=\"en\">", "<head>", "<meta charset=\"utf-8\">", "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">", 
                "<style>p\u007bline-height:1rem;\u007dtable\u007bborder-collapse:collapse;padding-right:2em;\u007dtr, td, th, table\u007bborder: 1px solid black;padding: 0.3rem 1rem 0.3rem 1rem;\u007d.subsection\u007b margin-left: 3em;\u007d.sectionend\u007bmargin-bottom:3rem;\u007d.indent\u007b margin-left: 3em;\u007d"
                "tr:nth-child(even)\u007bbackground-color:#cccccc;\u007dtr:nth-child(odd)\u007bbackground-color:#eeeeee;\u007dth\u007bbackground-color:#66dddd\u007d</style>", 
                "<title>Table Comparison</title>", "</head>", "<body style=\"font-size: 1.5rem;margin-right:2em;\">",
                f"<h1 id=\"top\"> Table Comparison - \"{filename1}\" and \"{filename2}\"</h1>",
                "<h2> File Information </h2>",
                "<div class=\"subsection\">",
                f"<p style=\"white-space:nowrap;\">File 1: {file1}</p>",
                f"<p style=\"white-space:nowrap;\">File 2: {file2}</p>",
                "</div>",
                "<h2> Quick Reference </h2>",
                "<div class=\"subsection\">",
                f"<p> Raw Columns in Table 1: {len(cols1)}</p>",
                f"<p> Raw Columns in Table 2: {len(cols2)}</p>",
                f"<p>Total Rows in Table 1: {len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT}</p>",
                f"<p>Total Rows in Table 2: {len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT}</p>",
                f"<p> Raw percent of rows in {filename1} that are in {filename2}: {round((raw_yes1 / (len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p> Raw percent of rows in {filename2} that are in {filename1}: {round((raw_yes2 / (len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p> Number of columns in common: {len(same_columns)}</p>",
                f"<p> Percent of rows in reordered {filename1} that are in reordered {filename2}: {round((swap_yes1 / (len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p> Percent of rows in reordered {filename2} that are in reordered {filename1}: {round((swap_yes2 / (len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                "</div>",   
                "<h2> Quick Links </h2>",
                "<a href=\"#cc1\">Columns in Table 1</a> |",
                "<a href=\"#cc2\">Columns in Table 2</a> |<br>",
                "<a href=\"#rrc\">Raw Row Comparison</a> |",
                "<a href=\"#rrc12\">Raw Row Comparison - Compare Table 1 with Table 2</a> |",
                "<a href=\"#rrc21\">Raw Row Comparison - Compare Table 2 with Table 1</a> |<br>",
                "<a href=\"#rrcc\">Reordered Row Comparison</a> |",
                "<a href=\"#rrcc12\">Reordered Row Comparison - Compare Table 1 with Table 2</a> |",
                "<a href=\"#rrcc21\">Reordered Row Comparison - Compare Table 2 with Table 1</a> |<br>",
                "<a href=\"#cvd\">Column Value Discrepancies</a><br><br>",
                "<h2> Column Comparison </h2>",
                "<a href=\"#top\">Back to Top</a>",
                "<div class=\"subsection\">",
                f"<p id=\"cc1\">Columns in Table 1: {len(cols1)}</p>",
                "<table>",
                "".join([f"<th>{i}</th>" for i in cols1]),
                "</table>",
                f"<p id=\"cc2\">Columns in Table 2: {len(cols2)}</p>",
                "<table>",
                "".join([f"<th>{i}</th>" for i in cols2]),
                "</table>",
                "</div>",
                "<h2 id=\"rrc\"> Raw Row Comparison </h2>",
                "<div class=\"subsection\">",
                "<a href=\"#top\">Back to Top</a>",
                f"<p>Total Rows in Table 1: {(len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT)}</p>",
                f"<p class=\"sectionend\">Total Rows in Table 2: {(len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT)}</p>",

                f"<h3 id=\"rrc12\"> Comparing {filename1} to {filename2}</h3>",
                "<a href=\"#top\">Back to Top</a>",
                f"<p class=\"indent\">Number of rows from {filename1} that appear in {filename2} -- {raw_yes1}</p>",
                f"<p class=\"indent\">Number of rows from {filename1} that do not appear in {filename2} -- {raw_no1}</p>",
                f"<p class=\"indent\"> Percent of rows in {filename1} that are in {filename2}: {round((raw_yes1 / (len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from {filename1} that do not appear in {filename2}</p>",
                "<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in cols1]),
                "</tr>",
                "".join(["<tr>" + "".join([f"<td>{rows1[i][j]}</td>" for j in range(len(rows1[i]))]) + "</tr>" for i in raw_idx1]),
                "</table>",

                f"<h3 id=\"rrc21\"> Comparing {filename2} to {filename1}</h3>",
                "<a href=\"#top\">Back to Top</a>",
                f"<p  class=\"indent\">Number of rows from {filename2} that appear in {filename1} -- {raw_yes2}</p>",
                f"<p  class=\"indent\">Number of rows from {filename2} that do not appear in {filename1} -- {raw_no2}</p>",
                f"<p class=\"indent\"> Percent of rows in {filename2} that are in {filename1}: {round((raw_yes2 / (len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from {filename2} that do not appear in {filename1}</p>",
                "<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in cols2]),
                "</tr>",
                "".join(["<tr>" + "".join([f"<td>{rows2[i][j]}</td>" for j in range(len(rows2[i]))]) + "</tr>" for i in raw_idx2]),
                "</table></div>",


                "<h2 id=\"rrcc\"> Reordered Row / Column Comparison </h2>",
                "<div class=\"subsection\">",
                "<a href=\"#top\">Back to Top</a>",
                "<h3> Summary </h3>",
                f"<p class=\"indent\"> Number of columns in Table 1 that do no appear in Table 2: {len(col1_cp)}</p>",
                f"<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in col1_cp]),
                "</tr></table>",
                f"<p class=\"indent\"> Number of columns in Table 2 that do no appear in Table 1: {len(col2_cp)}</p>",
                f"<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in col2_cp]),
                "</tr></table>",
                f"<p class=\"indent\"> Number of columns that appear in both tables: {len(same_columns)}</p>",
                f"<p class=\"indent\"> Columns that appear in both tables:</p><table class=\"indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr></table>",
                f"<p class=\"indent\">Total Rows in Table 1: {len(doc1)}</p>",
                f"<p class=\"sectionend indent\">Total Rows in Table 2: {len(doc2)}</p>",

                f"<h3 id=\"rrcc12\"> Comparing reordered {filename1} to reordered {filename2}</h3>",
                "<a href=\"#top\">Back to Top</a>",
                f"<p class=\"indent\">Number of rows from reordered {filename1} that appear in reordered {filename2} -- {swap_yes1}</p>",
                f"<p class=\"indent\">Number of rows from reordered {filename1} that do not appear in reordered {filename2} -- {swap_no1}</p>",
                f"<p class=\"indent\"> Percent of rows in reordered {filename1} that are in reordered {filename2}: {round((swap_yes1 / (len(doc1) if len(doc1) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from reordered {filename1} that do not appear in reordered {filename2}</p>",
                "<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                "".join(["<tr>" + "".join([f"<td>{rows1_swap[i][j]}</td>" for j in range(len(rows1_swap[i]))]) + "</tr>" for i in swap_idx1]),
                "</table>",

                f"<h3 id=\"rrcc21\">Comparing reordered {filename2} to reordered {filename1} </h3>",
                "<a href=\"#top\">Back to Top</a>",
                f"<p class=\"indent\" >Number of rows from reordered {filename2} that appear in reordered {filename1} -- {swap_yes2}</p>",
                f"<p class=\"indent\" >Number of rows from reordered {filename2} that do not appear in reordered {filename1} -- {swap_no2}</p>",
                f"<p class=\"indent\"> Percent of rows in reordered {filename2} that are in reordered {filename1}: {round((swap_yes2 / (len(doc2) if len(doc2) <= LINE_LIMIT else LINE_LIMIT)) * 100, 2)}%</p>",
                f"<p class=\"indent\"> Rows from reordered {filename2} that do not appear in reordered {filename1}</p>",
                "<table class=\"sectionend indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                "".join(["<tr>" + "".join([f"<td>{rows2_swap[i][j]}</td>" for j in range(len(rows2_swap[i]))]) + "</tr>" for i in swap_idx2]),
                "</table></div>",
                "<h2 id=\"cvd\"> Column Value Discrepancies using value 1 / value 2 (Has false positives if rows are not in the same order in source tables)</h2>",
                "<a href=\"#top\">Back to Top</a>",
                "<table class=\"indent\"><tr>",
                "".join([f"<th>{i}</th>" for i in same_columns]),
                "</tr>",
                "".join(["<tr>" + "".join([f"<td> {discrepancies[j][i]} </td>" for j in same_columns]) + "</tr>" for i in range(len(discrepancies[same_columns[0]]))]),
                "<table style=\"user-select:none;-webkit-user-select:none;-ms-user-select:none;margin-left: 6em;opacity: 0;\"><tr>",
                "".join([f"<th>{i}</th>" for i in cols1])+"<th>space</th>",
                "</tr></table>",
                "</body>", 
                "</html>"
            ]

    filename = f"TableComparison_{filename1}_{filename2}"
    count = 1
    while os.path.exists(filename + ".html"):
        filename = f"TableComparison_{filename1}_{filename2}_{count}"
        count += 1

    file = open(filename + ".html", "w", encoding="utf-8") 
    file.write("\n".join(output))
    print(f"Table Comparison Report saved as {filename}.html")
    print("HTML report has been generated.")
    file.close()

    try:
        webbrowser.open(os.path.abspath(filename+'.html'))
    except Exception as e:
        messagebox.showerror("Error Opening File", "An error occurred when attempting to display the report. To open the report manually, locate the file \"" + os.path.abspath(filename+'.html') + "\"")



# for i in range(len(files)):
#     file1 = files[i]
#     file2 = files2[i]
#     main()

if __name__ == "__main__":
    main()
