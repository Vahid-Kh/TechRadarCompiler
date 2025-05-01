Tech Radar Compiler
The Tech Radar Compiler is a tool designed to generate radar visualizations from input data stored in Excel files. The radar visualizations are automatically opened in your predefined browser and saved as HTML files.

Support and Code Modifications
For support and code modifications, please contact:

Stefano Menengello: Python compiler support
Vahid Khorshidi: HTML code & EXE code generators support
Executable File Creation
To create an executable file (.exe) from the Python script, follow these steps:

Open a Python terminal and install the necessary packages:

pip install pyinstaller
pip install pandas
pip install xlrd
Create the executable file using PyInstaller:

pyinstaller --onefile -w html2py.py
Place the html2py.exe file together with the TechRadarInputs.xlsx in the same directory.

Run html2py.exe and enjoy the radar visualization.

Usage Instructions
Save the Excel file with input data in a local directory and specify its path in the filename variable within the script.

Specify the name and version of the radar.

The radar will automatically open in your predefined browser and save itself as an HTML file in the same path where the .py file is stored.

Note: If you call two radars with the same name, the most recent will overwrite the older one.

Inputs
Filename: Path to the Excel file containing input data.
Version: Specify the version of the radar.
Example
# Example input configuration
filename = r'Business unit 1/TechRadarInputs.xlsx'
version = "V2.0"
License
This project is licensed under the MIT License. See the LICENSE file for more information.
