![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Qt](https://img.shields.io/badge/Qt-%23217346.svg?style=for-the-badge&logo=Qt&logoColor=white)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
# Data Management Application

This is a Python application built with PyQt5 that allows users to manage data in an Excel file. The application provides a graphical user interface (GUI) where users can insert, update, and delete data in the Excel file. It also supports different themes, such as dark mode and light mode, to enhance the user experience.

### data-entry-viewer-form

# Installation
To run the application, you need to have Python and PyQt5 installed on your system. You can install the dependencies using the following command:
```
pip install PyQt5 openpyxl pandas qdarktheme
```
# Usage
1. Run the application by executing the following command:
```
python main.py
```
2. GUI Interface: The application opens a graphical user interface (GUI) where you can perform various operations on the data.
*Insert Data: Enter the required information in the input fields and click the "Insert" button to add a new row to the table and Excel file.
* Update Data: Double-click on a cell in the table to edit its contents. Press Enter or click outside the cell to save the changes to the table and Excel file.
* Delete Data: Select one or more rows in the table and click the "Delete" button to remove them from the table and Excel file.
* Theme Selection: Choose between dark mode and light mode by selecting the corresponding radio button.

3. File Operations: The application provides options to create a new file or open an existing file.
  * New File: Click on the "File" menu and select "New File" to create a new Excel file. Choose a location and file name for the new file.
  * Open File: Click on the "File" menu and select "Open File" to open an existing Excel file. Browse to the desired file and open it. The data from the file will be displayed in the table.

# Limitations

* The application currently supports only Excel files (.xlsx, .xls) for data storage. Other file formats are not supported.
* The input validation is limited to checking the name for letters only and the age for positive values. Additional validation or constraints can be added as per the specific requirements.

# License
This project is licensed under the _MIT License_.

Feel free to contribute, report issues, or provide suggestions for improvement.

# Acknowledgements

This application was built using the following libraries:
* PyQt5: A Python binding for the Qt framework, used for creating the GUI.
* openpyxl: A library for reading and writing Excel files, used for data management.
* pandas: A powerful data manipulation library, used for loading data from Excel files.
* qdarktheme: A library for applying dark themes to PyQt5 applications.
Special thanks to the developers of these libraries for their contributions.
