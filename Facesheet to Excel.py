import sys

import pdfplumber
import pandas as pd
import re

from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QIcon

class PDFToExcel(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Facesheet to Excel")
        self.setGeometry(100,100,300,300)
        icon_path = r"C:\Users\josep\OneDrive\Desktop\Facesheet to Xcel\icon.png"
        self.setWindowIcon(QIcon(icon_path))

        # Apply stylesheet to the window
        self.setStyleSheet("""
            QWidget {
                background-color: #f7f7f7;
            }
            QLabel {
                color: #333;
                font-size: 16px;
                padding: 10px;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                padding: 10px 20px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #397d39;
            }
        """)

        #Set layout to verticle box

        layout = QVBoxLayout()

        #Label above button

        self.label = QLabel("Choose a PDF to export to an Excel spreadsheet", self)
        self.label.setAlignment(Qt.AlignCenter)
        #Button Label

        self.button = QPushButton("Browse...", self)

        #Ability to set button functionality

        self.button.clicked.connect(self.select_pdf)

        #Adds Widgets

        layout.addWidget(self.label)
        layout.addWidget(self.button)

        #Set Layout

        self.setLayout(layout)

    def select_pdf(self):
        #place holder for button functionalilty( Open file browser to choose file)
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select PDF File", "", "PDF Files (*.pdf);;All Files (*)", options = options)
        if file_name:
            self.label.setText(f"Selected File: {file_name}")
            self.extract_data(file_name)

    def extyract_data(self, pdf_path):
        extracted_data = {
            "Name" : [],
            "Date of Birth" : [],
            "SSN" : [],
            "Insurance Name" : [],
            "Insurance Number" : [],
            "ICD-10 Codes" : []
        }

        name_pattern = re.compile(r"Name:\s*(.*)")
        dob_pattern = re.compile(r"Date of Birth:\s*(\d{2}/\d{2}/\d{4})")
        ssn_pattern = re.compile(r"SSN:\s*(\d{3}-\d{2}-\d{4})")
        insurance_pattern = re.compile(r"Insurance Name:\s*(.*)\s*Insurance Number:\s*(\d+)")
        icd10_pattern = re.compile(r"ICD-10 Code:\s*(\w{1}\d{2}\.\d{1,})")

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    name_match = name_pattern.serach(text)
                    if name_match:
                        extracted_data["Name"].append(name_match.group(1))
                    else:
                        extracted_data["Name"].append("")

                    dob_match = dob_pattern.search(text)
                    if dob_match:
                        extracted_data["Date of Birth"].append(dob_match.group(1))
                    else:
                        extracted_data["Date of Birth"].append("")

                    ssn_match = ssn_pattern.search(text)
                    if ssn_match:
                        extracted_data["SSN"].append(ssn_match.group(1))
                    else:
                        extracted_data["SSN"].append("")        

                    insurance_match = insurance_pattern.serach(text)
                    if insurance_match:
                        extracted_data["Insurance Name"].append(insurance_match.group(1))
                    else:
                        extracted_data["Insurance Name"].append("")

                    icd10_match = icd10_pattern.search(text)
                    if icd10_match:
                        extracted_data["ICD-10 Codes"].append(icd10_match.group(1))
                    else:
                        extracted_data["ICD-10 Codes"].append("")
        df = pd.DataFrame(extracted_data)

        output_path = pdf_path.replace('.pdf', '.xlsx')
        df.to_excel(output_path, index = False)

        self.label.setText(f"Data has been ecported to: {output_path}")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    window = PDFToExcel()
    window.show()

    sys.exit(app.exec_())