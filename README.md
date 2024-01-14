# Exclter

Exclter is a Python based program which was created to help with the filtering of multiple large tables in a CRM (Customer Relationship Management) project. Project included of handling large Excel files with alot of information about clients, those files needed to be filtered to only contain the more relevant information for the company using the CRM-service. This program was designed to be used for documents which are being imported to Pipedrive, but can also be used for other usecases regarding Excel files.

## Features

1. Excel File Processing: Exclter allows users to select an untreated Excel file and process it according to specific requirements. It supports both .xlsx and .xls formats.
2. Custom Column Selection: Users can choose which columns to retain in the processed file. Exclter supports the selection of columns either through a text file including a list of column names or by using a set of default columns predefined in the source code.
3. Data Filtering Options: Specialized filtering options are available, such as the ability to filter out certain rows based on specific criteria like 'Chairman' or 'CEO' titles.
4. Additional Data Fields: Users can add custom data fields, like 'Lead Title' and 'Industry', to the processed file, enhancing its utility.
5. Output Customization: Exclter provides the flexibility to specify the name and save location of the output file.
6. Data Verification and Splitting: Post-processing, the program verifies data consistency between the original and processed files.
7. Keeps the original files: The original table is kept unchanged.

## Installation

#### Prerequisites

Python 3.x and the necessary libraries (PyQt5 and Pandas) are required to run the program.

Install the libraries using the command:
```
pip install PyQt5 pandas
```

## Usage

To use Exclter:

1. Select Excel File: Select the untreated Excel file that you wish to process.
2. Choose Columns: Decide whether to use default columns or select a text file containing custom column names that you want to retain.
3. Filter Options: If needed, enable filtering based on specific titles like 'Chairman' or 'CEO'.
4. Add Additional Fields: Optionally, add additional data fields such as 'Lead Title' or 'Industry'.
5. Specify Output: Choose the name and location for the processed Excel file.
6. Process File: Click on 'Process File' to start the processing. The status and results will be displayed in the log section.
7. Data Verification and Splitting: After processing, the program will verify the consistency of the data.

## Future plans

1. Add more options to the process, eg. by specifying the filtering by column values. (Now CEO & Chairmen of the board only supported)
2. Option to add a custom named column with a custom value, now it is locked to only lead title and industry.
3. Add an option to split the final table to desired amount or keep the file in single table.
4. Beautify the GUI.