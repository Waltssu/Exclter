import sys
import os
from PyQt5.QtCore import pyqtSignal, pyqtSlot, Qt
from PyQt5.QtWidgets import QTextEdit, QCheckBox, QApplication, QMainWindow, QWidget, QLabel, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog, QMessageBox, QLineEdit
import pandas as pd 
import threading
import shutil


class MainWindow(QMainWindow):
    update_status_signal = pyqtSignal(str)
    processing_complete_signal = pyqtSignal()

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel File Processor")
        self.setGeometry(100, 100, 500, 200)
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout()
        self.central_widget.setLayout(self.layout)
        self.filter_chairmen = False
        self.filter_ceo = False

        # Initialize save folder path as None
        self.save_folder = None

        # Default columns
        self.default_columns = [
            # FILL DEFAULT COLUMNS HERE
        ]

        self.setup_ui()

        self.excel_file = None
        self.columns_text_file = None
        self.output_file = None

        # Connect signals
        self.update_status_signal.connect(self.update_status)
        self.processing_complete_signal.connect(self.processing_complete)

    def setup_ui(self):
        # Step 1: Select Untreated Excel Table
        self.step1_label = QLabel("1. Select Untreated Excel Table")
        self.step1_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.step1_label)

        self.step1_v_layout = QVBoxLayout()
        self.layout.addLayout(self.step1_v_layout)

        self.step1_button = QPushButton("Browse")
        self.step1_button.setFixedWidth(100)
        self.step1_button.clicked.connect(self.select_excel_table)
        self.step1_v_layout.addWidget(self.step1_button)

        self.step1_file_label = QLabel("No file selected")
        self.step1_v_layout.addWidget(self.step1_file_label) 

        # Step 2: Select Columns Text File
        self.step2_label = QLabel("2. Select Columns Text File")
        self.step2_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.step2_label)

        # Checkbox for using default columns
        self.use_defaults_checkbox = QCheckBox("Use Defaults")
        self.layout.addWidget(self.use_defaults_checkbox)
        self.use_defaults_checkbox.stateChanged.connect(self.toggle_defaults)

        self.step2_v_layout = QVBoxLayout()
        self.layout.addLayout(self.step2_v_layout)

        self.step2_button = QPushButton("Browse")
        self.step2_button.setFixedWidth(100)
        self.step2_button.clicked.connect(self.select_columns_text_file)
        self.step2_v_layout.addWidget(self.step2_button)

        self.step2_text_label = QLabel("No file selected")
        self.step2_v_layout.addWidget(self.step2_text_label)


        # Input Field: Lead title
        self.lead_title_label = QLabel("Lead title")
        self.lead_title_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.lead_title_label)

        self.lead_title_input = QLineEdit()
        self.lead_title_input.setPlaceholderText("Enter lead title here")
        self.layout.addWidget(self.lead_title_input)

        # Input Field: Whose industry?
        self.whose_industry_label = QLabel("Whose industry?")
        self.whose_industry_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.whose_industry_label)

        self.whose_industry_input = QLineEdit()
        self.whose_industry_input.setPlaceholderText("Enter industry here")
        self.layout.addWidget(self.whose_industry_input)
        
        # Modified Filter Checkbox
        self.chairman_ceo_filter_checkbox = QCheckBox("Chairmen and CEO not selected")
        self.layout.addWidget(self.chairman_ceo_filter_checkbox)
        self.chairman_ceo_filter_checkbox.stateChanged.connect(self.toggle_chairman_ceo_filter)

        # Set initial state of the checkbox to unchecked
        self.chairman_ceo_filter_checkbox.setChecked(True)     

        # Output File Name
        self.step3_label = QLabel("3. Output File Name")
        self.step3_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.step3_label)

        # Output Filename Input Field
        self.output_filename_input = QLineEdit()
        self.output_filename_input.setPlaceholderText("Enter output file name here")
        self.layout.addWidget(self.output_filename_input)

        # Select Save Folder
        self.save_folder_label = QLabel("Select Save Folder")
        self.save_folder_label.setStyleSheet("font-weight: bold;")
        self.layout.addWidget(self.save_folder_label)

        self.save_folder_input = QLineEdit(self)
        self.save_folder_input.setPlaceholderText("No folder selected")
        self.layout.addWidget(self.save_folder_input)

        # Layout for save and current buttons
        self.save_folder_buttons_layout = QHBoxLayout()

        # Browse-button
        self.save_folder_button = QPushButton("Browse")
        self.save_folder_button.setFixedWidth(120)
        self.save_folder_button.clicked.connect(self.select_save_folder)
        self.save_folder_buttons_layout.addWidget(self.save_folder_button, alignment=Qt.AlignLeft)

        # Use current-button
        self.use_current_button = QPushButton("Use Current Folder")
        self.use_current_button.setFixedWidth(120)
        self.use_current_button.clicked.connect(self.use_current_folder)
        self.save_folder_buttons_layout.addWidget(self.use_current_button, alignment=Qt.AlignLeft)

        # Add a stretch factor to push the buttons to the left
        self.save_folder_buttons_layout.addStretch(1)

        # Add the layout to the main layout
        self.layout.addLayout(self.save_folder_buttons_layout)

        # Process File Button - Centered
        self.process_button = QPushButton("Process File")
        self.process_button.setEnabled(False)
        self.process_button.clicked.connect(self.process_file)
        self.process_button.setFixedWidth(100)
        self.layout.addWidget(self.process_button, alignment=Qt.AlignCenter)

        # Status Label
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)  # Make it read-only
        self.log_text_edit.setFixedHeight(0)
        self.layout.addWidget(self.log_text_edit)

    # Modify your update_status method to update the log widget
    @pyqtSlot(str)
    def update_status(self, message):
        self.log_text_edit.append(message)  # Append new log messages

    @pyqtSlot()
    def processing_complete(self):
        self.process_button.setEnabled(False)
        self.update_status_signal.emit("Process completed!")
        self.update_status_signal.emit("You can close this window now.")

    def select_save_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder_path:
            self.save_folder_input.setText(folder_path)
            self.update_process_button()  # Update the process button state

    def select_excel_table(self):
        file_dialog = QFileDialog(self, "Select Untreated Excel Table")
        file_dialog.setNameFilter("Excel Files (*.xlsx *.xls)")
        if file_dialog.exec_():
            file_path = file_dialog.selectedFiles()[0]
            self.set_excel_table(file_path)
            self.step1_file_label.setText(os.path.basename(file_path))

    def set_excel_table(self, file_path):
        self.excel_file = file_path
        base_name = os.path.basename(file_path)
        name_part = base_name.split('.')[0].split()[0]
        default_output = f"{name_part}_filtered.xlsx"
        self.output_filename_input.setText(default_output)
        self.update_process_button()

    def toggle_chairman_ceo_filter(self, state):
        self.filter_chairmen = self.filter_ceo = state == Qt.Checked
        if state != Qt.Checked:
            self.chairman_ceo_filter_checkbox.setText("Chairmen and CEO not selected")
        else:
            self.chairman_ceo_filter_checkbox.setText("Chairmen and CEO selected")

    
    def toggle_defaults(self, state):
        if state == Qt.Checked:
            # Use the default columns and update the label
            self.columns_text_file = 'default'  # Indicates the use of default columns
            self.step2_text_label.setText("Default columns selected")
            self.step2_button.setDisabled(True)
        else:
            # Reset to allow user to select a custom file
            self.columns_text_file = None
            self.step2_text_label.setText("No file selected")
            self.step2_button.setDisabled(False)

        self.update_process_button()  # Update the process button state


    def select_columns_text_file(self):
        if not self.use_defaults_checkbox.isChecked():
            file_dialog = QFileDialog(self, "Select Columns Text File")
            file_dialog.setNameFilter("Text Files (*.txt)")
            file_dialog.fileSelected.connect(self.set_columns_text_file)
            if file_dialog.exec_():
                file_path = file_dialog.selectedFiles()[0]
                self.columns_text_file = file_path
                self.step2_text_label.setText(file_path)
                self.update_process_button()
        else:
            pass

    def set_columns_text_file(self, file_path):
        self.columns_text_file = file_path
        self.step2_text_label.setText(file_path)
        self.update_process_button()

    def use_current_folder(self):
        if self.excel_file:
            current_folder = os.path.dirname(self.excel_file)
            self.save_folder_input.setText(current_folder)
            self.update_process_button()  # Update the process button state

    def update_process_button(self):
        # The process button should be enabled only if a save folder is selected
        # and either the use defaults checkbox is checked or a columns file is selected
        if self.save_folder_input.text() and (self.use_defaults_checkbox.isChecked() or self.columns_text_file):
            self.process_button.setEnabled(True)
        else:
            self.process_button.setEnabled(False)

    def compare_dataframes(self, original_df, processed_df, sample_size=10):
        discrepancies = []

        # Get a sample of 'Kontaktin sähköpostiosoite' values from the processed dataframe
        sample_sposti = processed_df['Kontaktin sähköpostiosoite'].dropna().sample(n=sample_size, random_state=10)

        for sposti in sample_sposti:
            row_original = original_df[original_df['Kontaktin sähköpostiosoite'] == sposti]
            row_processed = processed_df[processed_df['Kontaktin sähköpostiosoite'] == sposti]

            if not row_original.empty and not row_processed.empty:
                # Iterate through each column in the original dataframe
                for column in original_df.columns:
                    # Find the matching column in the processed dataframe
                    if column in processed_df.columns:
                        original_value = row_original.iloc[0][column]
                        processed_value = row_processed.iloc[0][column]

                        # Compare the values
                        if pd.isna(original_value) and pd.isna(processed_value):
                            continue
                        if original_value != processed_value:
                            discrepancies.append(f"Discrepancy for {sposti} in column '{column}': Original = {original_value}, Processed = {processed_value}")

        return discrepancies


    def process_file(self):
        self.output_file = self.output_filename_input.text()
        if self.output_file:
            threading.Thread(target=self.process_file_task).start()
            self.process_button.setEnabled(False)
            self.update_status_signal.emit("Processing, please wait...")
        else:
            self.show_error_message("Output File Name is Empty", "Please specify an output file name.")

    def process_file_task(self):
        self.log_text_edit.setFixedHeight(250)
        self.update_status_signal.emit("Starting processing...")
        try:
            # Copy the Excel file to the save folder
            copied_file_path = os.path.join(self.save_folder_input.text(), self.output_filename_input.text())
            shutil.copy(self.excel_file, copied_file_path)
            self.update_status_signal.emit("File copied to the save folder!")

            # Check if default columns are to be used
            if self.columns_text_file == 'default':
                self.update_status_signal.emit("Default columns selected...")
                columns_to_keep = self.default_columns
            else:
                with open(self.columns_text_file, 'r') as file:
                    self.update_status_signal.emit("Not using default columns, reading custom file...")
                    columns_to_keep = file.read().splitlines()

            # Read the Excel file with all columns as strings
            df = pd.read_excel(copied_file_path, dtype=str)
            self.update_status_signal.emit("Excel file loaded...")            

            # Filter columns
            df = df[[col for col in columns_to_keep if col in df.columns]]
            self.update_status_signal.emit("Columns filtered...")

            # Filter rows based on 'Titteli' column
            if self.filter_chairmen or self.filter_ceo:
                conditions = []
                if self.filter_chairmen:
                    conditions.append(df['Titteli'] == "Hallituksen puheenjohtaja")
                if self.filter_ceo:
                    conditions.append(df['Titteli'] == "Toimitusjohtaja")
                df = df[pd.concat(conditions, axis=1).any(axis=1)]
                self.update_status_signal.emit("Rows filtered based on Titteli...")

            # Add new columns 'Lead title', 'Lead owner', 'Person owner', 'Organization owner'
            lead_title_value = self.lead_title_input.text()
            whose_industry_value = self.whose_industry_input.text()

            new_columns = {
                'Lead Title': lead_title_value,
                'Lead Owner': whose_industry_value,
                'Person Owner': whose_industry_value,
                'Organization Owner': whose_industry_value
            }

            # Normalize existing column names for comparison
            existing_columns_normalized = {col.lower().replace(' ', ''): col for col in df.columns}

            for col_name, col_value in new_columns.items():
                col_name_normalized = col_name.lower().replace(' ', '')
                if col_name_normalized not in existing_columns_normalized:
                    df[col_name] = col_value
                    self.update_status_signal.emit(f"Column '{col_name}' created!")
                else:
                    self.update_status_signal.emit(f"Column '{col_name}' already exists, will not overwrite!")


            # Save the processed file
            save_path = os.path.join(self.save_folder_input.text(), self.output_filename_input.text()) if self.save_folder_input.text() else self.output_file
            df.to_excel(save_path, index=False)
            self.update_status_signal.emit("Processed file saved!")

            # After processing, read both original and processed files
            self.update_status_signal.emit("Verifying the data for consistency...")
            original_df = pd.read_excel(self.excel_file)
            processed_df = pd.read_excel(os.path.join(self.save_folder_input.text(), self.output_filename_input.text()))

            # Compare the dataframes
            discrepancies = self.compare_dataframes(original_df, processed_df)
            if not discrepancies:
                self.update_status_signal.emit("Verification successful: Data is consistent!")
                self.update_status_signal.emit("Creating copies of the filtered file for splitting...")

                # Splitting the dataframe
                num_rows = len(processed_df)
                rows_per_split = num_rows // 4

                # Extracting the base name of the original file without extension
                base_name = os.path.basename(self.excel_file)
                name_part = os.path.splitext(base_name)[0]

                # Creating copies of the filtered file and splitting them
                for i in range(4):
                    copy_file_name = f"{name_part} part {i+1}-4.xlsx"
                    copy_file_path = os.path.join(self.save_folder_input.text(), copy_file_name)
                    shutil.copy(os.path.join(self.save_folder_input.text(), self.output_filename_input.text()), copy_file_path)
                    self.update_status_signal.emit(f"Copy {i+1} created and being split as '{copy_file_name}'")

                    # Read and split each copy
                    split_df = pd.read_excel(copy_file_path, dtype=str)
                    start_row = i * rows_per_split
                    end_row = start_row + rows_per_split if i < 3 else num_rows

                    # Adjusting the split dataframe
                    split_df = split_df.iloc[start_row:end_row].copy()

                    # Save the split back to the same file
                    split_df.to_excel(copy_file_path, index=False)
        
            else:
                discrepancy_messages = '\n'.join(discrepancies)
                self.update_status_signal.emit(f"Verification failed, data inconsistencies found in:\n{discrepancy_messages}")

            self.processing_complete_signal.emit()

        except Exception as e:
            self.update_status_signal.emit(f"An error occurred: {str(e)}")

    def show_error_message(self, title, message):
        QMessageBox.critical(self, title, message)

    def toggle_filter_flag(self):
        self.filter_chairmen = not self.filter_chairmen
        self.filter_button.setStyleSheet("background-color: green;" if self.filter_chairmen else "")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())