import os
import tkinter as tk
from tkinter import ttk, messagebox
import threading
import time  # For periodic update delay
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import re

# Import ttkbootstrap Style
from ttkbootstrap import Style

# Define Google Drive API scope
SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets']

labels_frame = None  # Define labels_frame globally
folder_file_count = {}  # Dictionary to store folder file counts

def authenticate():
    """Authenticate with Google APIs."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json')
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Construct the path to credentials.json dynamically
            script_dir = os.path.dirname(os.path.realpath(__file__))
            credentials_path = os.path.join(script_dir, 'credentials.json')
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def extract_numbers(filename):
    """Extract numbers from the file name using regex."""
    regex = r'[a-zA-Z\s]*\s*(\d+)\s*(?:\(\d+\))?\s*\.pdf'
    match = re.search(regex, filename)
    if match:
        return match.group(1)
    else:
        return None

def select_folders(tree):
    """Select folders from Google Drive."""
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    folders = service.files().list(
        q="mimeType='application/vnd.google-apps.folder'",
        fields="files(id, name)").execute().get('files', [])
    selected_folders = []
    if folders:
        popup = tk.Toplevel()
        popup.title("Select Folders")
        popup.geometry("400x300")
        popup.attributes('-topmost', True)

        # Create a canvas with a scrollbar
        canvas = tk.Canvas(popup)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(popup, orient=tk.VERTICAL, command=canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame inside the canvas
        frame = tk.Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor=tk.NW)

        # Function to update canvas scrollregion
        def configure_scrollregion(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        frame.bind("<Configure>", configure_scrollregion)

        # Function to fetch files data
        def fetch_files():
            nonlocal selected_folders
            index = 1  # Initialize index variable
            for folder_name in selected_folders:
                folder_query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
                folder = service.files().list(q=folder_query, fields="files(id)").execute().get('files', [])
                if folder:
                    folder_id = folder[0]['id']
                    files = service.files().list(q=f"'{folder_id}' in parents", fields="files(id, name, parents, webViewLink)").execute().get('files', [])
                    for file in files:
                        file_name = file['name']
                        index_str = extract_numbers(file_name)
                        
                        # Fetch parent folder name
                        parent_folder_id = file.get('parents', [])[0]
                        parent_folder = service.files().get(fileId=parent_folder_id, fields="name").execute().get('name', '')
                        
                        url = file.get('webViewLink', '')
                        tree.insert("", "end", values=("", index, "", index_str, file_name, parent_folder, url))
                        index += 1  # Increment index for each file
                    
                    # Update folder_file_count dictionary
                    folder_file_count[folder_name] = len(files)

                    # Print folder name and file count to console
                    print_to_console(f"{folder_name} : {len(files)}")

        # Function to get selected folders
        def get_selected_folders():
            nonlocal selected_folders
            for child, var in checkbutton_value.items():
                if var.get() == 1:  # Check if the checkbutton is selected
                    selected_folders.append(child.cget("text"))
            popup.destroy()
            # Start background fetch thread
            background_fetch_thread = BackgroundFetchThread(tree, selected_folders)
            background_fetch_thread.start()

        # Create checkbuttons for folders
        checkbutton_value = {}  # Initialize checkbutton value dictionary
        for folder in folders:
            var = tk.IntVar(value=0)
            checkbutton = tk.Checkbutton(frame, text=folder['name'], variable=var)
            checkbutton.pack(anchor=tk.W)
            checkbutton_value[checkbutton] = var  # Store checkbutton value

        # Create select button
        select_button = tk.Button(popup, text="Select", command=get_selected_folders)
        select_button.pack(pady=10)

class BackgroundFetchThread(threading.Thread):
    def __init__(self, tree, selected_folders):
        super().__init__()
        self.tree = tree
        self.selected_folders = selected_folders

    def run(self):
        creds = authenticate()
        service = build('drive', 'v3', credentials=creds)
        index = 1  # Initialize index variable
        chunk_size = 10  # Number of files to fetch in each chunk
        for folder_name in self.selected_folders:
            folder_query = f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'"
            folder = service.files().list(q=folder_query, fields="files(id)").execute().get('files', [])
            if folder:
                folder_id = folder[0]['id']
                page_token = None
                while True:
                    files = service.files().list(q=f"'{folder_id}' in parents", fields="nextPageToken, files(id, name, parents, webViewLink)", pageToken=page_token).execute()
                    page_token = files.get('nextPageToken')
                    files = files.get('files', [])
                    for file in files:
                        file_name = file['name']
                        index_str = extract_numbers(file_name)
                        
                        # Fetch parent folder name
                        parent_folder_id = file.get('parents', [])[0]
                        parent_folder = service.files().get(fileId=parent_folder_id, fields="name").execute().get('name', '')
                        
                        url = file.get('webViewLink', '')
                        self.tree.insert("", "end", values=("", index, "", index_str, file_name, parent_folder,"", url))
                        index += 1  # Increment index for each file
                    if not page_token:
                        break
                    time.sleep(2)  # Add a delay between each page fetch

def list_google_sheets(service):
    """List all Google Sheets in Google Drive."""
    results = service.files().list(
        q="mimeType='application/vnd.google-apps.spreadsheet'",
        fields="files(id, name)").execute()
    sheets = results.get('files', [])
    return sheets

def list_tabs(service, sheet_id):
    """List tabs of a Google Sheet."""
    sheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet.get('sheets', [])
    tabs = [sheet['properties']['title'] for sheet in sheets]
    return tabs

def column_to_letter(column):
    """Convert column number to letter."""
    letters = ''
    while column > 0:
        column, remainder = divmod(column - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters

def list_columns(service, sheet_id, tab_name):
    """List columns of a Google Sheets tab."""
    # Get the spreadsheet
    sheet = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheets = sheet.get('sheets', [])

    # Find the specified tab by name
    tab = None
    for s in sheets:
        if s['properties']['title'] == tab_name:
            tab = s
            break

    if not tab:
        messagebox.showerror("Error", f"Tab '{tab_name}' not found in the Google Sheet.")
        return []

    # Get the total number of columns in the tab
    total_columns = tab['properties']['gridProperties']['columnCount']

    # Construct the range from column A to the last column
    range_name = f"{tab_name}!A1:{column_to_letter(total_columns)}1"

    # Retrieve values from the specified range to get column names
    result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    columns = values[0] if values else []

    # Create a list of tuples containing column name and its letter identifier
    column_info = [(col_name, column_to_letter(idx + 1)) for idx, col_name in enumerate(columns)]

    return column_info

service = None  # Define a global variable to store the service object
sheet_id = None  # Define a global variable to store the sheet ID

def select_sheet(tree):
    """Select Google Sheet from a list."""
    global service  # Access the global service variable
    global sheet_id  # Access the global sheet_id variable
    creds = authenticate()
    service = build('drive', 'v3', credentials=creds)
    sheets = list_google_sheets(service)
    if sheets:
        sheet_names = [sheet['name'] for sheet in sheets]
        sheet_list_window = tk.Toplevel()
        sheet_list_window.title("Select Google Sheet")
        sheet_list_window.geometry("300x200")
        selected_sheet = tk.StringVar(value=sheet_names[0])
        sheet_listbox = tk.Listbox(sheet_list_window, listvariable=selected_sheet, selectmode="single")
        sheet_listbox.pack(expand=True, fill="both", side="left")

        # Add scrollbar for the sheet listbox
        sheet_scrollbar = tk.Scrollbar(sheet_list_window, orient="vertical")
        sheet_scrollbar.pack(side="right", fill="y")
        sheet_scrollbar.config(command=sheet_listbox.yview)
        sheet_listbox.config(yscrollcommand=sheet_scrollbar.set)

        for sheet_name in sheet_names:
            sheet_listbox.insert(tk.END, sheet_name)

        def on_ok(tree):
            global sheet_id # Access the outer scope variable
            selected_sheet_name = sheet_listbox.get(tk.ACTIVE)
            print("Selected Sheet:", selected_sheet_name)  # Print selected sheet to console
            # Fetch the selected sheet and its ID
            sheet = [sheet for sheet in sheets if sheet['name'] == selected_sheet_name][0]
            sheet_id = sheet['id']
            print("Sheet ID:", sheet_id)  # Print sheet ID to console

            global service  # Access the global service variable
            service = build('sheets', 'v4', credentials=creds)  # Use Google Sheets API
            tabs = list_tabs(service, sheet_id)
            if tabs:
                select_tab_window = tk.Toplevel()
                select_tab_window.title("Select Tab")
                select_tab_window.geometry("300x200")
                selected_tab = tk.StringVar(value=tabs[0])
                tab_listbox = tk.Listbox(select_tab_window, listvariable=selected_tab, selectmode="single")
                tab_listbox.pack(expand=True, fill="both", side="left")

                # Add scrollbar for the tab listbox
                tab_scrollbar = tk.Scrollbar(select_tab_window, orient="vertical")
                tab_scrollbar.pack(side="right", fill="y")
                tab_scrollbar.config(command=tab_listbox.yview)
                tab_listbox.config(yscrollcommand=tab_scrollbar.set)

                for tab_name in tabs:
                    tab_listbox.insert(tk.END, tab_name)

                def on_tab_ok():
                    selected_tab_name = tab_listbox.get(tk.ACTIVE)
                    print("Selected Tab:", selected_tab_name)  # Print selected tab to console
                    # Now let's list the columns for the selected tab
                    columns = list_columns(service, sheet_id, selected_tab_name)
                    if columns:
                        select_column_window = tk.Toplevel()
                        select_column_window.title("Select Columns")
                        select_column_window.geometry("300x200")

                        # Add scrollbar for the column listbox
                        column_scrollbar = tk.Scrollbar(select_column_window, orient="vertical")
                        column_scrollbar.pack(side="right", fill="y")

                        def on_column_ok():
                            selected_columns = {'IDs': None, 'Phone number': None}
                            for col_name, col_letter, col_var in column_vars:
                                if col_var.get():
                                    if selected_columns['IDs'] is None:
                                        selected_columns['IDs'] = col_letter
                                    elif selected_columns['Phone number'] is None:
                                        selected_columns['Phone number'] = col_letter
                                    else:
                                        messagebox.showerror("Error", "Please select only two columns.")
                            print("Selected Columns:", selected_columns)
                            select_column_window.destroy()
                            select_tab_window.destroy()
                            sheet_list_window.destroy()
                            matching_thread = MatchingValuesThread(service, sheet_id, selected_tab_name, selected_columns['IDs'], tree)
                            matching_thread.start()

                        column_vars = []
                        column_listbox = tk.Listbox(select_column_window, yscrollcommand=column_scrollbar.set)
                        column_listbox.pack(expand=True, fill="both", side="left")

                        column_scrollbar.config(command=column_listbox.yview)

                        for column_info in columns:
                            col_name = column_info[0]
                            col_letter = column_info[1]
                            col_var = tk.BooleanVar()
                            col_var.set(False)
                            column_vars.append((col_name, col_letter, col_var))
                            checkbox = tk.Checkbutton(column_listbox, text=f"{col_name} ({col_letter})", variable=col_var)
                            checkbox.pack(anchor="w")

                        column_ok_button = tk.Button(select_column_window, text="OK", command=on_column_ok)
                        column_ok_button.pack(side="top", pady=5)
                    else:
                        messagebox.showerror("Error", "No columns found in the selected Google Sheets tab.")

                tab_ok_button = tk.Button(select_tab_window, text="OK", command=on_tab_ok)
                tab_ok_button.pack()

            else:
                messagebox.showerror("Error", "No tabs found in the selected Google Sheet.")

        ok_button = tk.Button(sheet_list_window, text="OK", command=lambda: on_ok(tree))
        ok_button.pack()

    else:
        messagebox.showerror("Error", "No Google Sheets found in Google Drive.")



def periodic_update(tree):
    """Periodic update function."""
    # Implement code to update data in the ttk.Treeview periodically
  
    while True:
        time.sleep(5)
        #Sprint("Periodic update")

def print_to_console(text):
    """Print text to both console and label."""
    print(text)
    

def main():
    global service  # Access the global service variable
    global sheet_id  # Access the global sheet_id variable
    global tree  # Access the global tree variable
    root = tk.Tk()
    root.title("Google Drive File Selection")

    # Apply ttkbootstrap theme
    style = Style(theme='minty')
    style.configure('TButton', font=('Helvetica', 12))

    tree = ttk.Treeview(root, columns=("Check", "NO",  "GS-name", "Index","File Name", "Folder Name","GS-Column", "URL"), show="headings")

    # Set column headings with left alignment
    tree.heading("Check", text="Check", anchor="w")
    tree.heading("NO", text="NO", anchor="w")
    tree.heading("GS-name", text="GS-name", anchor="w")
    tree.heading("Index", text="Index", anchor="w")
    tree.heading("File Name", text="File Name", anchor="w")
    tree.heading("Folder Name", text="Folder Name", anchor="w")
    tree.heading("GS-Column", text="GS-Column", anchor="w")
    tree.heading("URL", text="URL", anchor="w")

    tree.column("Check", width=50)
    tree.column("NO", width=50)
    tree.column("GS-name", width=100)
    tree.column("Index", width=100)
    tree.column("File Name", width=200)
    tree.column("Folder Name", width=200)
    tree.column("GS-Column", width=100)
    tree.column("URL", width=400)

    # Set column alignments
    tree.column("Check", anchor="center")
    tree.column("NO", anchor="center")
    tree.column("GS-name", anchor="w")
    tree.column("Index", anchor="center")
    tree.column("File Name", anchor="w")
    tree.column("Folder Name", anchor="w")
    tree.column("GS-Column", anchor="w")
    tree.column("URL", anchor="w")

    tree.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")

    # Create vertical scrollbar
    yscrollbar = ttk.Scrollbar(root, orient='vertical', command=tree.yview)
    yscrollbar.grid(row=1, column=1, sticky='ns')
    tree.configure(yscrollcommand=yscrollbar.set)
    
# Create the labeled frame
    menu_frame = ttk.LabelFrame(root, text="Menu")
    menu_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky='we')

    # Create and position the buttons within the labeled frame
    select_files_button = ttk.Button(menu_frame, text="Select Folders", command=lambda: select_folders(tree))
    select_files_button.grid(row=0, column=1, padx=(10, 0), pady=10, sticky='w')

    select_sheet_button = ttk.Button(menu_frame, text="Select Sheet", command=lambda: select_sheet(tree))
    select_sheet_button.grid(row=0, column=2, padx=10, pady=10)

    link_url_button = ttk.Button(menu_frame, text="Link URL", command=start_extract_thread)
    link_url_button.grid(row=0, column=3, padx=(0, 10), pady=10, sticky='e')


    clear_tree_button = ttk.Button(menu_frame, text="Clear Tree", command=lambda: clear_tree(tree))
    clear_tree_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')
    # Start periodic update thread
    periodic_update_thread = threading.Thread(target=periodic_update, args=(tree,))
    periodic_update_thread.daemon = True  # Daemonize the thread
    periodic_update_thread.start()

    # Configure grid weights to allow expansion
    root.columnconfigure(0, weight=1)
    root.rowconfigure(1, weight=1)

    # Create a frame for labels
    labels_frame = tk.LabelFrame(root, text="Summary")
    labels_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky="ew")
    tree.bind("<Double-1>", lambda event: on_double_click(event, root, tree))
    root.mainloop()
def clear_tree(tree):
    """Clear all items in the ttk.Treeview."""
    for item in tree.get_children():
        tree.delete(item)
def on_double_click(event, root, tree):
    """Copy the URL to the clipboard when double-clicked."""
    item = tree.identify('item', event.x, event.y)  # Identify the item clicked
    if item:
        values = tree.item(item, "values")
        url = values[7] if len(values) > 7 else None  # Assuming 'URL' is at index 7 in the values array
        if url:
            root.clipboard_clear()  # Clear the clipboard
            root.clipboard_append(url)  # Copy the URL to the clipboard
            print("URL copied to clipboard:", url)
        else:
            print("No URL available for the selected item.")

# Bind double-click event to the tree


def fetch_google_sheet_data(service, sheet_id, tab_name, column_letter):
    """Fetch data from the specified column of the Google Sheet."""
    range_name = f"{tab_name}!{column_letter}:{column_letter}"
    result = service.spreadsheets().values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    
    # Extract values from the fetched data
    column_data = {}
    for row_number, row in enumerate(values, start=1):
        if row:
            item = row[0]
            cell_reference = f"{column_letter}{row_number}"  # Construct cell reference
            
            if item:
                try:
                    item = int(item)  # Convert to integer
                    column_data[item] = cell_reference  # Add item as key and cell_reference as value to the dictionary
                except ValueError:
                    print(f"Warning: Unable to convert {item} to an integer.")
    
    # Print fetched values for debugging
    print(f"Selected column {column_letter}: {column_data}")
    print(f"Total number of items in column {column_letter}: {len(column_data)}")
    
    return column_data



    # Assuming 'tree' is a ttk tree widget instance
"""def extract_items(tree):
    #Extract the GS-Column item and its corresponding URL item.
    extracted_items = {}
    for item in tree.get_children():
        values = tree.item(item, "values")
        gs_column = values[6]  # Assuming 'GS-Column' is at index 6 in the values array
        url = values[7]  # Assuming 'URL' is at index 7 in the values array
        extracted_items[gs_column] = url
    print("Extracted Items:", extracted_items)  # Print the extracted dictionary
    return extracted_items"""
"""def extract_items(tree, service, sheet_id):
    #Extract the GS-Column item and its corresponding URL item.
    
    def paste_values_to_sheet(extracted_items):
        #Paste values into the Google Sheet
        for cell_reference, url in extracted_items.items():
            range_name = f"{cell_reference}:{cell_reference}"
            value_input_option = 'RAW'
            value_range_body = {
                'values': [[url]]
            }
            
            try:
                result = service.spreadsheets().values().update(
                    spreadsheetId=sheet_id, range=range_name,
                    valueInputOption=value_input_option, body=value_range_body).execute()
                print(f"URL '{url}' pasted to cell {cell_reference} successfully.")
            except Exception as e:
                print(f"Error occurred while pasting URL '{url}' to cell {cell_reference}: {str(e)}")
    
    # Main extraction logic
    extracted_items = {}
    for item in tree.get_children():
        values = tree.item(item, "values")
        gs_column = values[6]  # Assuming 'GS-Column' is at index 6 in the values array
        url = values[7]  # Assuming 'URL' is at index 7 in the values array
        extracted_items[gs_column] = url

    # Call the nested function to paste values into the Google Sheet
    paste_values_to_sheet(extracted_items)

    print("Extracted Items:", extracted_items)  # Print the extracted dictionary
    return extracted_items"""

def start_extract_thread():
    extract_thread = ExtractItemsThread(tree, service, sheet_id)
    extract_thread.start()

class ExtractItemsThread(threading.Thread):
    def __init__(self, tree, service, sheet_id):
        super().__init__()
        self.tree = tree
        self.service = service
        self.sheet_id = sheet_id

    def run(self):
        extracted_items = self.extract_items()
        self.paste_values_to_sheet(extracted_items)

    def extract_items(self):
        extracted_items = {}
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            gs_column = values[6]  # Assuming 'GS-Column' is at index 6 in the values array
            url = values[7]  # Assuming 'URL' is at index 7 in the values array
            extracted_items[gs_column] = url
        return extracted_items

    def paste_values_to_sheet(self, extracted_items):
        for cell_reference, url in extracted_items.items():
            if cell_reference:
                range_name = f"{cell_reference}:{cell_reference}"
                value_input_option = 'RAW'
                value_range_body = {
                    'values': [[url]]
                }
                try:
                    result = self.service.spreadsheets().values().update(
                        spreadsheetId=self.sheet_id, range=range_name,
                        valueInputOption=value_input_option, body=value_range_body).execute()
                    print(f"URL '{url}' pasted to cell {cell_reference} successfully.")
                    self.update_checkmark(cell_reference, "✔️")
                except Exception as e:
                    print(f"Error occurred while pasting URL '{url}' to cell {cell_reference}: {str(e)}")
                    self.update_checkmark(cell_reference, "❌")

    def update_checkmark(self, cell_reference, checkmark):
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            gs_column = values[6]  # Assuming 'GS-Column' is at index 6 in the values array
            if gs_column == cell_reference:
                self.tree.item(item, values=(checkmark, *values[1:]))  # Only update the "Check" column
                break


class MatchingValuesThread(threading.Thread):
    def __init__(self, service, sheet_id, tab_name, column_letter, tree):
        super().__init__()
        self.service = service
        self.sheet_id = sheet_id
        self.tab_name = tab_name
        self.column_letter = column_letter
        self.tree = tree
        print("Matching thread starting ...")
        
    def run(self):
        # Extract index column values from the ttk tree
        index_column_values = self.extract_index_column_values()
        
        # Fetch data from the Google Sheet
        column_data = fetch_google_sheet_data(self.service, self.sheet_id, self.tab_name, self.column_letter)
        
        if column_data:
            # Compare values and print matching values with their corresponding dictionary
            matched_values = self.compare_and_print_matching_values(column_data)
            if matched_values:
                cell_references, values = zip(*matched_values.items())
                # Create the popup window
                self.create_link_popup(cell_references, values, matched_values)
            else:
                print("No matching values found.")
        else:
            print("No data fetched from Google Sheet.")
    
    def create_link_popup(self, cell_references, values, matched_values):
        """Create a popup window for entering the link column."""
        # Create a popup window
        popup_window = tk.Toplevel()
        popup_window.title("Enter Link Column")

        # Add a label to the popup window
        label = tk.Label(popup_window, text="Enter the link column (e.g., 'T'): ")
        label.pack()

        # Add an entry widget to the popup window
        entry = tk.Entry(popup_window)
        entry.pack()

        # Add an OK button to confirm the input
        ok_button = tk.Button(popup_window, text="OK", command=lambda: self.on_ok(entry, popup_window, cell_references, values, matched_values))
        ok_button.pack()

    def on_ok(self, entry, popup_window, cell_references, values, matched_values):
        """Handle OK button click in the link column popup."""
        new_column = entry.get().strip().upper()  # Get the entered column value
        popup_window.destroy()  # Close the popup window

        # Create a new dictionary with updated cell references
        updated_matched_values = {}
        for value, cell_reference in matched_values.items():
            updated_cell_reference = new_column + cell_reference[1:]
            updated_matched_values[value] = updated_cell_reference

        print("Updated Cell References:")
        for value, updated_cell_reference in updated_matched_values.items():
            print(f"Value: {value}, Cell Reference: {updated_cell_reference}")

        # Iterate through the ttk tree to update GS-Column for corresponding rows
        for row_item in self.tree.get_children():
            row_values = self.tree.item(row_item, "values")
            index_value = row_values[3] if len(row_values) > 3 else None
            if index_value is not None and int(index_value) in updated_matched_values:
                old_reference = updated_matched_values[int(index_value)]
                row_number = int(old_reference[1:])  # Extract the row number from the old reference
                updated_reference = new_column + str(row_number)  # Append the new column to the row number
                self.tree.set(row_item, "GS-Column", updated_reference)

    def extract_index_column_values(self):
        """Extract index column values from the ttk tree."""
        index_column_values = []
        for item in self.tree.get_children():
            values = self.tree.item(item, "values")
            # Assume the index value is at index 3 in the values array
            index_value = values[3] if len(values) > 3 else None
            if index_value is not None:
                try:
                    index_value_int = int(index_value)  # Convert to integer
                    index_column_values.append(index_value_int)  # Append index value to the array
                except ValueError:
                    print(f"Warning: Unable to convert index value '{index_value}' to an integer.")
        return index_column_values

    def compare_and_print_matching_values(self, column_data):
        """Compare values from the Google Sheet column data with index column values and update the GS-name column in the ttk tree."""
        print("column data in the thread:", column_data)
        matched_values = {}
        if column_data:
            print("Matching Values with Dictionary:")
            children = self.tree.get_children()
            for i in range(len(children)):  # Iterate over all items in the ttk tree
                item = children[i]
                try:
                    value = int(self.tree.item(item, "values")[3])
                except ValueError:
                    print("Warning: Unable to convert index value to an integer.")
                    continue  # Skip this item and proceed to the next one
                # Extract value from the 'Index' column ********(assuming it's the 4th column)
                cell_reference = column_data.get(value)  # Get the cell reference from the column data based on the index value
                if cell_reference is not None:
                    print(f"Value: {value}, Cell Reference: {cell_reference}")
                    self.tree.set(item, "GS-name", cell_reference)  # Update GS-name column in the ttk tree
                    matched_values[value] = cell_reference
        else:
            print("No data fetched from Google Sheet.")
        return matched_values

    def update_tree(self, cell_references, values, matched_values, new_column):
        """Update the ttk tree with the new cell references."""
        for item, cell_reference in zip(self.tree.get_children(), cell_references):
            # Assume 'GS-name' is the column index for the 'GS-name' column
            # Convert cell_reference to string to perform subscript operation
            cell_reference_str = str(cell_reference)
            # Ensure cell_reference_str has at least two characters
            if len(cell_reference_str) >= 2:
                updated_cell_reference = new_column + cell_reference_str[1:]
                self.tree.set(item, "GS-Column", updated_cell_reference)
            else:
                print(f"Error: Invalid cell reference format for {cell_reference}")

# Create an instance of the MatchingValuesThread class

if __name__ == "__main__":
    main()

