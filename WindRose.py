# Import necessary libraries
import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from windrose import WindroseAxes
import numpy as np
import openpyxl
import os
import xml.etree.ElementTree as ET

def read_data(file_path):
    # Read the file and find the start of the data
    with open(file_path, 'r') as file:
        lines = file.readlines()
        data_start = next(i for i, line in enumerate(lines) if line.startswith('%|           date'))
    
    # Read the CSV data, skipping the header rows
    global data
    data = pd.read_csv(file_path, skiprows=data_start+1, 
                       names=['date', 'WindSpeed_10m', 'WindDirection_10m', 'HourlyRainfall', 
                              'AirTemperature', 'CloudCover', 'SolarRadiation', 'RelativeHumidity', 
                              'BarometricPressure', 'NegAirtemp'])
    
    # Clean the data
    data['WindSpeed_10m'] = pd.to_numeric(data['WindSpeed_10m'], errors='coerce')
    data['WindDirection_10m'] = pd.to_numeric(data['WindDirection_10m'], errors='coerce')
    
    # Remove rows with NaN or infinite values in wind speed or direction
    data = data.replace([np.inf, -np.inf], np.nan).dropna(subset=['WindSpeed_10m', 'WindDirection_10m'])
    
    # Ensure wind direction is between 0 and 360
    data['WindDirection_10m'] = data['WindDirection_10m'] % 360
    
    return data

def create_wind_rose(data):
    # Extract wind speed and direction data
    wind_speed = data['WindSpeed_10m']
    wind_dir = data['WindDirection_10m']
    
    # Create a wind rose plot
    ax = WindroseAxes.from_ax()
    
    # Plot the wind rose using a bar chart
    ax.bar(wind_dir, wind_speed, normed=True, opening=0.8, edgecolor='white',
           bins=[0, 2, 5, 10, 15, 50])
    
    # Set the legend for wind speed ranges
    ax.set_legend(title="Wind Speed (m/s)", labels=['0-2', '2-5', '5-10', '10-15', '15+'])
    
    # Set the title of the wind rose plot
    ax.set_title(f"Wind Rose - {os.path.basename(file_path)}")
    
    # Return the figure object containing the wind rose plot
    return ax.figure

def create_frequency_table(data):
    # Define bins for wind speed and direction
    speed_bins = [0, 2, 5, 10, 15, 50]
    speed_labels = ['0-2', '2-5', '5-10', '10-15', '15+']
    dir_bins = np.arange(-11.25, 360, 22.5)
    dir_labels = [f"{i:.1f}" for i in np.arange(0, 360, 22.5)]

    # Categorize wind speed and direction data
    data['speed_bin'] = pd.cut(data['WindSpeed_10m'], bins=speed_bins, labels=speed_labels, include_lowest=True)
    data['dir_bin'] = pd.cut(data['WindDirection_10m'], bins=dir_bins, labels=dir_labels, include_lowest=True)

    # Create frequency table
    freq_table = pd.crosstab(data['speed_bin'], data['dir_bin'], normalize=True)
    return freq_table

def import_file():
    # Open file dialog to select a file
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("DLP files", "*.dlp"), ("Text files", "*.txt")])
    if file_path:
        # Read the data from the selected file
        data = read_data(file_path)
        
        # Create and display frequency table
        freq_table = create_frequency_table(data)
        for widget in left_frame.winfo_children():
            widget.destroy()
        tree = ttk.Treeview(left_frame)
        tree["columns"] = ["Wind Speed (m/s)"] + list(freq_table.columns)
        tree["show"] = "headings"
        for column in tree["columns"]:
            tree.heading(column, text=column)
            tree.column(column, width=50, anchor="center")
        tree.column("Wind Speed (m/s)", width=100, anchor="w")
        for index, row in freq_table.iterrows():
            values = [index] + [f"{val:.4f}" for val in row]
            tree.insert("", "end", values=values)
        tree.pack(fill=tk.BOTH, expand=True)
        
        # Create and display wind rose
        fig = create_wind_rose(data)
        for widget in right_frame.winfo_children():
            widget.destroy()
        canvas = FigureCanvasTkAgg(fig, master=right_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

def save_wind_rose():
    # Save wind rose figure as jpg
    fig = plt.gcf()
    file_path = filedialog.asksaveasfilename(defaultextension=".jpg", filetypes=[("JPEG files", "*.jpg")])
    if file_path:
        fig.savefig(file_path, format='jpg', dpi=300, bbox_inches='tight')

def save_table():
    freq_table = create_frequency_table(data)
    file_path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML files", "*.xml")])
    if file_path:
        root = ET.Element("Data")
        ET.SubElement(root, "Information").text = "No information"
        ET.SubElement(root, "Name").text = "MyWindRose"
        
        # Define velocity bands
        velocity_bands = "2 5 10 15 50"
        ET.SubElement(root, "Velocity_Bands").text = velocity_bands
        
        # Create Headings_Probabilities
        headings_probabilities = ET.SubElement(root, "Headings_Probabilities")
        
        for column in freq_table.columns:
            heading_prob = ET.SubElement(headings_probabilities, "Heading_Probabilities")
            heading_prob.text = " ".join(f"{val:.4f}" for val in freq_table[column])
        
        tree = ET.ElementTree(root)
        tree.write(file_path, xml_declaration=True, encoding='utf-8', method="xml")

# Create main window
root = tk.Tk()
root.title("Wind Data Viewer")
root.geometry("1500x1000")

# Create top frame for buttons
top_frame = tk.Frame(root)
top_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=10)

# Create import button
import_button = tk.Button(top_frame, text="Import File", command=import_file)
import_button.pack(side=tk.LEFT, padx=5)

# Create save wind rose button
save_wind_rose_button = tk.Button(top_frame, text="Save Wind Rose as JPG", command=save_wind_rose)
save_wind_rose_button.pack(side=tk.LEFT, padx=5)

# Create save table button
save_table_button = tk.Button(top_frame, text="Save Table as MERIT XML", command=save_table)
save_table_button.pack(side=tk.LEFT, padx=5)

# Create left frame for frequency table
left_frame = tk.Frame(root, width=400)
left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=10, pady=10)

# Create right frame for wind rose
right_frame = tk.Frame(root)
right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=10, pady=10)

# Start the main event loop
root.mainloop()