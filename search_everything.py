import os
import tkinter as tk
from tkinter import filedialog, ttk
import subprocess
from datetime import datetime

# Function to convert bytes to a human-readable size (KB/MB/GB)
def format_size(size):
    if size >= 1_073_741_824:  # If file is larger than 1GB
        return f"{size / 1_073_741_824:.2f} GB"
    elif size >= 1_048_576:  # If file is larger than 1MB
        return f"{size / 1_048_576:.2f} MB"
    elif size >= 1_024:  # If file is larger than 1KB
        return f"{size / 1_024:.2f} KB"
    else:
        return f"{size} bytes"  # For files smaller than 1KB

# Function to search for files/folders in the directory
def search_files(directory, search_term, wildcard=True):
    results = []
    for root, dirs, files in os.walk(directory):  # Traverse directories recursively
        for name in dirs + files:
            if wildcard:  # Wildcard search (case-insensitive)
                if search_term.lower() in name.lower() and not name.lower().endswith('.lnk'):
                    path = os.path.join(root, name)
                    try:
                        size = os.path.getsize(path)
                        last_modified = os.path.getmtime(path)
                        results.append((name, path, size, last_modified))
                    except (OSError, FileNotFoundError):
                        continue
            else:  # Exact string search
                if search_term.lower() == name.lower() and not name.lower().endswith('.lnk'):
                    path = os.path.join(root, name)
                    try:
                        size = os.path.getsize(path)
                        last_modified = os.path.getmtime(path)
                        results.append((name, path, size, last_modified))
                    except (OSError, FileNotFoundError):
                        continue
    return results

# Function to open the selected file/folder in Windows Explorer
def open_in_explorer(path):
    subprocess.Popen(f'explorer /select,"{path.replace("/", "\\")}"')

# Create the GUI window
def open_gui():
    global root
    root = tk.Tk()
    root.title("Mini Desktop Google © NCCC IT")
    root.state('zoomed')  # Fullscreen by default
    root.config(bg="#f0f0f0")

    # Directory to Search field
    frame = tk.Frame(root, bg="#f0f0f0")
    frame.pack(pady=20, padx=20)
    
    heading_label = tk.Label(root, text="Mini Desktop Google © NCCC IT", bg="#f0f0f0", font=("verdana", 15, "bold"))
    heading_label.place(relx=0.01, rely=0.02, anchor="w")
    
    dir_label = tk.Label(frame, text="Directory to Search:", bg="#f0f0f0", font=("verdana", 8))
    dir_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

    dir_entry = tk.Entry(frame, width=30, font=("verdana", 8))
    dir_entry.grid(row=0, column=1, padx=10, pady=5)
    dir_entry.insert(0, "C:/")

    def browse_directory():
        selected_directory = filedialog.askdirectory(initialdir="C:/", title="Select a Directory")
        if selected_directory:
            dir_entry.delete(0, tk.END)
            dir_entry.insert(0, selected_directory)

    browse_button = tk.Button(frame, text="Browse", font=("verdana", 8), command=browse_directory)
    browse_button.grid(row=0, column=2, padx=10, pady=5)

    term_label = tk.Label(frame, text="Search Term:", bg="#f0f0f0", font=("verdana", 8))
    term_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")

    term_entry = tk.Entry(frame, width=30, font=("verdana", 8))
    term_entry.grid(row=1, column=1, padx=10, pady=5)

    wildcard_var = tk.BooleanVar(value=True)
    wildcard_checkbox = tk.Checkbutton(frame, text="Not sure of exact name", variable=wildcard_var, bg="#f0f0f0", font=("verdana", 8))
    wildcard_checkbox.grid(row=2, columnspan=3, pady=10)

    # Search button
    search_button = tk.Button(root, text="Search", font=("verdana", 10, "bold"), bg="#4CAF50", fg="white", command=lambda: on_search())
    search_button.pack(pady=10, side=tk.TOP, anchor="center")

    # Treeview for displaying results in columns
    result_frame = tk.Frame(root, bg="#f0f0f0")
    result_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    treeview = ttk.Treeview(result_frame, columns=("Name", "Path", "Size", "Date Modified"), show="headings")
    treeview.pack(fill=tk.BOTH, expand=True)

    # Define column headings and their properties
    treeview.heading("Name", text="Name", command=lambda: sort_by_column(0))  # Sort by Name
    treeview.heading("Path", text="Path", command=lambda: sort_by_column(1))  # Sort by Path
    treeview.heading("Size", text="Size", command=lambda: sort_by_column(2))  # Sort by Size
    treeview.heading("Date Modified", text="Date Modified", command=lambda: sort_by_column(3))  # Sort by Date Modified

    treeview.column("Name", width=200)
    treeview.column("Path", width=400)
    treeview.column("Size", width=100)
    treeview.column("Date Modified", width=150)

    # Initialize a dictionary to track the sort order for each column
    sort_order = {
        0: True,  # True means ascending, False means descending
        1: True,
        2: True,
        3: True
    }

    # Function to handle search
    def on_search(event=None):
        directory = dir_entry.get()
        search_term = term_entry.get()
        wildcard = wildcard_var.get()

        if not search_term:
            return

        results = search_files(directory, search_term, wildcard)

        for result in results:
            name, path, size, last_modified = result
            formatted_size = format_size(size)
            formatted_time = datetime.fromtimestamp(last_modified).strftime('%Y-%m-%d %H:%M:%S')
            treeview.insert("", tk.END, values=(name, path, formatted_size, formatted_time))

    # Function to close the window
    def close_window(event=None):
        root.quit()

    # Function to handle sorting by column
    def sort_by_column(col_index):
        # Toggle the sort order
        sort_order[col_index] = not sort_order[col_index]

        # Fetch current data from Treeview
        data = [(treeview.item(item)["values"], item) for item in treeview.get_children("")]

        # Determine the sort order based on the toggle state
        sorted_data = sorted(data, key=lambda x: x[0][col_index], reverse=not sort_order[col_index])

        # Clear the current items and insert sorted items
        for item in treeview.get_children():
            treeview.delete(item)

        for item in sorted_data:
            treeview.insert("", tk.END, values=item[0])

    # Function to open the selected file/folder on double-click
    def on_item_double_click(event):
        selected_item = treeview.selection()[0]  # Get the selected item (row)
        path = treeview.item(selected_item, "values")[1]  # Extract path (column 1)
        open_in_explorer(path)  # Open the file/folder location

    # Bind double-click event for opening the file/folder
    treeview.bind("<Double-1>", on_item_double_click)

    root.bind('<Return>', on_search)
    root.bind('<Escape>', close_window)
    root.mainloop()
    


if __name__ == "__main__":
    open_gui()
