import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create a dictionary of file names and corresponding sheet names
file_dict = {
    'RTWGame1.bcs': 'Game1',
    'RTWGame2.bcs': 'Game2',
    'RTWGame3.bcs': 'Game3',
    'RTWGame4.bcs': 'Game4',
    'RTWGame5.bcs': 'Game5',
    'RTWGame6.bcs': 'Game6',
    'RTWGame7.bcs': 'Game7',
    'RTWGame8.bcs': 'Game8',
    'RTWGame9.bcs': 'Game9'
}

# Prompt the user to enter a game directory
game_dir = input("Enter your game directory: ")

# Prompt the user to choose an option from Game1 to Game9
game_num = input("Choose an option from Game1 to Game9: ")

# Define the standard structure
middle_path = "Save/"

# Construct the final path and file name
final_path = os.path.join(game_dir, middle_path, f"Game{game_num}", f"RTW{game_num}.bcs")

# Read the text file into a pandas DataFrame
df = pd.read_csv(final_path, header=None)

# Get the corresponding sheet name from the dictionary
sheet_name = file_dict[file_alias]

# Export the DataFrame to the specified sheet in an Excel workbook
with pd.ExcelWriter('output.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)