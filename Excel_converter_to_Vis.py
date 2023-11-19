import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import tkinter as tk
from tkinter import filedialog

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

# Function to process and visualize the data
def process_and_visualize(file_path):
    # Load the data from the "Apartments" sheet
    apartments_data = pd.read_excel(file_path, sheet_name='Apartments')

    # Mapping 'yes' to 1 and 'no' (or NaN) to 0 for amenities
    amenities_columns = ['Allows Pets', 'Has a gym', 'Has a Pool', 'Has parking']
    apartments_data[amenities_columns] = apartments_data[amenities_columns].applymap(lambda x: 1 if x == 'yes' else 0)

    # Aggregating the average rent for each complex
    average_rent = apartments_data.groupby('Complex')['Monthly Rent'].mean().reset_index()

    # Aggregating the total number of amenities for each complex
    total_amenities = apartments_data.groupby('Complex')[amenities_columns].sum().reset_index()
    total_amenities['Total Amenities'] = total_amenities[amenities_columns].sum(axis=1)

    # Merging the two aggregated datasets
    complex_summary = pd.merge(average_rent, total_amenities[['Complex', 'Total Amenities']], on='Complex')

    # Creating the visualization
    sns.set(style="whitegrid")
    fig, ax1 = plt.subplots(figsize=(10, 6))
    sns.barplot(x='Complex', y='Monthly Rent', data=complex_summary, ax=ax1, color='b', label='Average Monthly Rent')
    ax2 = ax1.twinx()
    sns.lineplot(x='Complex', y='Total Amenities', data=complex_summary, ax=ax2, color='r', marker='o', label='Total Amenities')
    ax1.set_xlabel('Housing Complex')
    ax1.set_ylabel('Average Monthly Rent ($)', color='b')
    ax2.set_ylabel('Total Amenities', color='r')
    plt.title('Comparison of Housing Complexes by Average Rent and Total Amenities')
    ax1.legend(loc='upper left')
    ax2.legend(loc='upper right')
    plt.show()

# Main script
if __name__ == "__main__":
    file_path = load_excel_file()
    if file_path:
        process_and_visualize(file_path)
