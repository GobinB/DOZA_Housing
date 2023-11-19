import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog

# Function to load the Excel file
def load_excel_file():
    # Setting up the tkinter root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open a dialog to select the file
    file_path = filedialog.askopenfilename(
        title="Select an Excel file", 
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    return file_path

# Function to ask user preferences
def ask_user_preferences():
    preferences = {}
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Asking for user preferences
    preferences['Allows Pets'] = simpledialog.askstring("Preferences", "Allows Pets? (yes/no)")
    preferences['Has a gym'] = simpledialog.askstring("Preferences", "Has a gym? (yes/no)")
    preferences['Has a Pool'] = simpledialog.askstring("Preferences", "Has a Pool? (yes/no)")
    preferences['Has parking'] = simpledialog.askstring("Preferences", "Has parking? (yes/no)")
    preferences['Beds'] = simpledialog.askinteger("Preferences", "Number of Beds:")
    preferences['Max Rent'] = simpledialog.askinteger("Preferences", "Maximum Monthly Rent:")

    return preferences

# Function to filter and display the data
def filter_and_display(file_path, preferences):
    # Load the data from the "Apartments" sheet
    apartments_data = pd.read_excel(file_path, sheet_name='Apartments')

    # Filtering the data based on user preferences
    for amenity in ['Allows Pets', 'Has a gym', 'Has a Pool', 'Has parking']:
        if preferences[amenity].lower() == 'yes':
            apartments_data = apartments_data[apartments_data[amenity] == 'yes']
    
    apartments_data = apartments_data[apartments_data['Beds'] >= preferences['Beds']]
    apartments_data = apartments_data[apartments_data['Monthly Rent'] <= preferences['Max Rent']]

    # Display the filtered results
    if not apartments_data.empty:
        print("Matching Apartments:")
        print(apartments_data)
    else:
        print("No matching apartments found.")

# Main script
if __name__ == "__main__":
    file_path = load_excel_file()
    if file_path:
        preferences = ask_user_preferences()
        filter_and_display(file_path, preferences)
