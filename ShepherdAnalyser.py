import os
import tkinter as tk
import pytesseract
import tempfile
from numpy import average
from pandas import DataFrame
from PIL import Image
from tkinter import ttk, filedialog
from xml.etree import ElementTree as ET

class ShepherdAnalyser:
    def __init__(self, root):
        self.root = root
        self.root.title("Exceedance Analyser")
        self.root.geometry("650x400")

        self.mode = tk.StringVar()
        self.mode_options = ["Risk", "Overpressure Exceedance", "Impulse Exceedance", "Thermal Exceedance"]
        self.xml_file_path = None
        pytesseract.pytesseract.tesseract_cmd = "C:/Users/Faiq Raedaya/AppData/Local/Programs/Tesseract-OCR/tesseract.exe"
        #pytesseract.pytesseract.tesseract_cmd = r'Tesseract-OCR\tesseract.exe'
        
        current_folder = os.path.dirname(os.path.abspath(__file__))
        #image_path = os.path.join(current_folder, 'MES.png')
        #self.logo_image = tk.PhotoImage(file=image_path)
        #self.logo_label = tk.Label(root, image=self.logo_image)
        #self.logo_label.grid(row=0, column=0, padx=10, pady=10, sticky="nw")
        
        self.title = tk.Label(root, text="Shell Shepherd Output Analyser", font=('Arial', 20))
        self.title.grid(row=0, column=1, columnspan=2, pady=10, sticky="w")

        self.create_label("1. Select analysis mode:", 1, 1, 10, "w")
        self.mode_select = ttk.OptionMenu(root, self.mode, "Select mode", *self.mode_options)
        self.mode_select.grid(row=1, column=2, columnspan=2, pady=10, sticky="w")

        self.create_label("2. Select XML file:", 2, 1, 10, "w")
        self.import_button = self.create_button("Import", self.import_xml, 2, 2, 2, "w")

        self.create_label("3. Start analysis:", 3, 1, 10, "w")
        self.run_button = self.create_button("Run", self.analyse_xml, 3, 2, 2, "w")

        self.create_label("4. Export to Excel:", 4, 1, 10, "w")
        self.export_button = self.create_button("Export", self.export_to_excel, 4, 2, 2, "w")

        self.info_label = tk.Label(root, text="Please select analysis mode and import a file.", wraplength=600)
        self.info_label.grid(row=5, column=0, columnspan=3, pady=10)
        self.loading_bar = ttk.Progressbar(root, length=500, mode="determinate")
        self.loading_bar.grid(row=6, column=0, columnspan=3, pady=10)

    def create_label(self, text, row, column, pady, sticky):
        '''
        Create a tkinter label
        '''
        label = tk.Label(self.root, text=text)
        label.grid(row=row, column=column, pady=pady, sticky=sticky)

    def create_button(self, text, command, row, column, columnspan, sticky):
        '''
        Create a tkinter button
        '''
        button = tk.Button(self.root, text=text, command=command, state="active")
        button.grid(row=row, column=column, columnspan=columnspan, pady=10, sticky=sticky)
        return button

    def display_info(self, text):
        '''
        Updates tkinter info label text.
        '''
        self.info_label.config(text=text)

    def import_xml(self):
        '''
        Imports the .xml file using tkinter filedialog.
        '''
        self.xml_file_path = filedialog.askopenfilename(filetypes=[("XML files", "*.xml")])
        if self.xml_file_path:
            self.display_info(f"Imported {self.xml_file_path}.")
        else:
            self.display_info("Import cancelled.")

    def analyse_xml(self):
        '''
        Calls the appropriate function based on selected analysis mode.
        '''
        self.data = {}
        if self.xml_file_path:
            if "Risk" in self.mode.get():
                self.analyse_risk()
            elif "Exceedance" in self.mode.get():
                self.analyse_exceedance()
            else:
                self.display_info("Error: No mode selected.")
        else:
            self.display_info("Error: No file imported.")

    def analyse_risk(self):
        '''
        Analyses risk values in the .xml file.
        '''
        # Parse XML file using EventTree
        tree = ET.parse(self.xml_file_path)
        root = tree.getroot()

        # Find and enumerate all buildings, setup loading bar
        ns = {'ns': ''}
        buildings = root.findall(".//ns:OBJECT", namespaces=ns)
        buildings_total = len(buildings)
        buildings_processed = 0
        if buildings_total == 0:
            self.display_info("Error: No buildings found.")
        self.loading_bar["maximum"] = buildings_total
        self.loading_bar["value"] = buildings_processed

        # Iterate through all buildings
        for building_object in buildings:
            # Assign building name from object heading
            building_name = building_object.get("HEADING")

            # Find all risk results and initialise values
            results = building_object.find(".//ns:RESULTS", namespaces=ns)
            risk_sums = {"GAS_HITS": 0, "FLAME_HITS": 0, "THERMAL_HITS": 0, "CLOUD_FIRE_HITS": 0, "BLEVE_HITS": 0, "TOXIC_HITS": 0}
            risk_index = {"GAS_HITS": 3, "FLAME_HITS": 3, "THERMAL_HITS": 4, "CLOUD_FIRE_HITS": 3, "BLEVE_HITS": 2, "TOXIC_HITS": 4}

            # Iterate through all risk types
            for risk_type in risk_sums.keys():
                # Find all rows of risk type
                leak_rows = results.findall(f".//ns:{risk_type}/ns:LEAK_ROW", namespaces=ns)
                
                # Sum all values of risk type based on predefined column index
                leak_item_index = risk_index.get(risk_type, 4)
                risk_sum = sum(float(leak_row.find(f".//ns:LEAK_ITEM[{leak_item_index}]", namespaces=ns).text) for leak_row in leak_rows)
                
                # Special sum condition for thermal hits 
                risk_sum = 0 if risk_type == "THERMAL_HITS" and risk_sum > 0.01 else risk_sum
                
                # Store sum of risk type 
                risk_sums[risk_type] = risk_sum
            
            # Store all risk values of the building, update loading bar
            self.data[building_name] = {"Building Name": building_name, **risk_sums}
            buildings_processed += 1
            self.loading_bar["value"] = buildings_processed
            self.display_info(f"Processing {building_name}.\n Processed {buildings_processed} out of {buildings_total} buildings ({buildings_processed/buildings_total*100:.0f}%).")
            self.root.update()
        
        self.display_info(f"Analysis complete.\n Processed {buildings_total} buildings.")

    def analyse_exceedance(self):
        '''
        Analyses risk values in the .xml file.
        '''
        # Parse XML file using EventTree
        tree = ET.parse(self.xml_file_path)
        root = tree.getroot()

        # Initialise const and dict
        self.get_units()
        row_keys_index = ["1E-2", "1E-3", "1E-4", "1E-5", "1E-6", "1E-7", "1E-8", "1E-9"]
        row_keys = [f"{self.mode.get()} at {key} ({self.units})" for key in row_keys_index]
        row_values = [33, 70, 108, 145, 182, 219, 257, 293]
        row_range = range(min(row_values),max(row_values)+1)
        results = column_number = interpolated_value = dict(zip(row_keys, [0] * len(row_keys)))
        x_lower = 88 if self.mode.get() == "Thermal Exceedance" else 70
        x_upper = 636

        # Find and enumerate all buildings, setup loading bar
        ns = {'ns': ''}
        buildings = root.findall(".//ns:OBJECT", namespaces=ns)
        buildings_total = len(buildings)
        buildings_processed = 0
        if buildings_total == 0:
            self.display_info("Error: No buildings found.")
            return
        self.loading_bar["maximum"] = buildings_total
        self.loading_bar["value"] = buildings_processed

        # Iterate through all buildings
        for building_object in buildings:
            # Assign building name from object heading
            building_name = building_object.get("HEADING")

            # Call the array conversion function
            self.get_result_array(building_object, ns)

            # Report zeroes for all frequencies if the image does not exist
            if self.array is None:
                self.data[building_name] = {"Building Name": building_name, **{row_keys[i]: 0 for i in range(len(row_values))}}
                continue
            
            # Iterate through all rows of the graph
            for row in row_range:
                # Call the column number finder function
                column_number[row] = self.find_column(self.array, x_min=x_lower, x_max=x_upper, row_number=row)
                
                # Call the interpolation function
                interpolated_value[row] = self.interpolate_value(column_number[row], x_min=x_lower, x_max=x_upper, y_min=0, y_max=self.x_max)

                # Overwrite new value with the old value if it is lesser
                if (row != min(row_range)) and (interpolated_value[row] < interpolated_value[row-1]):
                    interpolated_value[row] = interpolated_value[row-1]
                
                # Store the result of each row
                results[row] = interpolated_value[row]
            
            # Store all exceedance values of the building, update loading bar
            self.data[building_name] = {"Building Name": building_name, **{row_keys[i]: results[row_values[i]] for i in range(len(row_values))}}
            buildings_processed += 1
            self.loading_bar["value"] = buildings_processed
            self.display_info(f"Processing {building_name}.\n Processed {buildings_processed} out of {buildings_total} buildings ({buildings_processed/buildings_total*100:.0f}%).")
            self.root.update()
        
        self.display_info(f"Analysis complete.\n Processed {buildings_total} buildings.")

    def get_units(self):
        '''
        Sets units based on the selected analysis mode
        '''
        if self.mode.get() == "Overpressure Exceedance":
            self.units = "psi"
        elif self.mode.get() == "Impulse Exceedance":
            self.units = "psi-ms"
        elif self.mode.get() == "Thermal Exceedance":
            self.units = "kW/m2"
        else:
            self.display_info("Error: Invalid mode.")

    def get_result_array(self, object, ns):
        '''
        Obtains the array 
        '''
        self.array = None
        # Get the image path using image path finder function based on selected analysis mode
        if self.mode.get() == "Overpressure Exceedance":
            self.image_path = self.get_image_path(self.xml_file_path, object, "GRAPH")
        elif self.mode.get() == "Impulse Exceedance":
            self.image_path = self.get_image_path(self.xml_file_path, object, "IMPULSE")
        elif self.mode.get() == "Thermal Exceedance":
            thermal_element = object.find(".//ns:RESULTS/ns:THERMAL_EXCEEDANCE", namespaces=ns)
            if thermal_element is None:
                return
            self.image_path = os.path.join(os.path.dirname(self.xml_file_path), thermal_element.text)
        else:
            self.display_info("Error: Invalid mode.")
            return

        if self.image_path is not None:
            # Calls image cropping function to get the maximum x-value of the graph
            self.cropped_image = self.crop_bottom_right(self.image_path)

            # Calls OCR function to read the maximum x-value of the graph
            self.x_max = self.convert_image_to_number(self.cropped_image)

            # Calls the image-to-array conversion function
            self.array = self.convert_image_to_array(self.image_path)

            # Remove the temporarily-stored cropped image
            os.remove(self.cropped_image)

    def get_image_path(self, xml_file_path, object, graph_name):
        '''
        Obtains the image path for a given 
        '''
        element = object.find(".//" + graph_name)
        if element is not None and element.text is not None:
            return os.path.join(os.path.dirname(xml_file_path), element.text)
        else:
            return None

    def crop_bottom_right(self, image_path):
        '''
        Crop the bottom right 40x40 of the image (max x-value of the graph) 
        '''
        # Open image with PIL
        original_image = Image.open(image_path)
        
        # Find image dimensions
        width, height = original_image.size

        # Define crop dimensions
        left = width - 44
        upper = height - 42
        right = width
        lower = height
        
        # Crop image
        cropped_image = original_image.crop((left, upper, right, lower))
        
        # Save cropped image as a temporary file
        temp_image_path = tempfile.NamedTemporaryFile(delete=False, suffix=".png").name
        cropped_image.save(temp_image_path)

        return temp_image_path
    
    def convert_image_to_number(self, image_path):
        '''
        Uses PyTesseractOCR to convert an image to a number
        '''
        # Open image with PIL
        img = Image.open(image_path)

        # Call pytesseract image-to-string conversion function
        text = pytesseract.image_to_string(img, config='--psm 6')

        # Return a number if valid
        try:
            number = int(text)
            return number
        except ValueError:
            return 0
        
    def convert_image_to_array(self, image_path):
        '''
        Converts an LxW image to an MxN array of RGB values  
        '''
        # Open image with PIL
        img = Image.open(image_path)
        # Find image dimensions
        width, height = img.size

        # Iterate through all y values
        array = []
        for y in range(height):
            row = []
            # Iterate through all x values
            for x in range(width):
                # Get the RGB value of the pixel at (x,y)
                pixel = img.getpixel((x, y))
                row.append(pixel)
            array.append(row)
        return array
    
    def find_column(self, array, x_min, x_max, row_number):
        '''
        Finds the column number of the first pixel to not be white/gray
        '''
        # Iterate throguh all columns
        for col_number in range(x_min, x_max):
            # Get the RGB values of the pixel
            pixel = array[row_number][col_number]
            
            # Return the column number if the pixel has a R/G value < 100
            if pixel[1] < 100 and pixel[2] < 100:
                return col_number
        return 0

    def interpolate_value(self, column, x_min, x_max, y_min, y_max):
        '''
        Simple interpolation function
        '''
        if column == 0:
            interpolated_value = 0
        else:
            interpolated_value = y_min + ((y_max - y_min) / (x_max - x_min)) * (column - x_min)
        return interpolated_value
        
    def export_to_excel(self):
        '''
        Exports a dict to an excel file
        '''
        # Call pandas dataframe orientation function
        df = DataFrame.from_dict(self.data, orient='index')

        # Prompt to save the file
        excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if excel_file_path:
            # Call pandas Excel conversion function
            df.to_excel(excel_file_path, index=False)
            self.display_info(f"Data exported to Excel file:\n{excel_file_path}")
        else:
            self.display_info(f"Error: Something went wrong with the export.\n Make sure the file you are trying to export to is not currently open.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ShepherdAnalyser(root)
    root.mainloop()