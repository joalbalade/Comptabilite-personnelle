import sys
import pandas as pd
import sqlite3
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QMessageBox


class CsvToSqliteApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Excel to SQLite')

        layout = QVBoxLayout()

        self.button_select_file = QPushButton('Select Excel File', self)
        self.button_select_file.clicked.connect(self.select_file)
        layout.addWidget(self.button_select_file)

        self.button_import = QPushButton('Import to SQLite', self)
        self.button_import.clicked.connect(self.import_to_sqlite)
        layout.addWidget(self.button_import)

        self.setLayout(layout)
        self.db_name = 'data.db'
        self.excel_file_path = ''

    def select_file(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)", options=options)
        if file_path:
            self.excel_file_path = file_path

    def import_to_sqlite(self):
        if not self.excel_file_path:
            QMessageBox.warning(self, 'Warning', 'Please select an Excel file first!')
            return

        try:
            conn = sqlite3.connect(self.db_name)
            cursor = conn.cursor()

            # Create table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS transactions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT,
                    libelle TEXT,
                    montant REAL
                )
            ''')

            # Skip the first few lines that are not part of the table
            df = pd.read_excel(self.excel_file_path, engine='openpyxl', skiprows=6)

            # Normalize and rename columns
            df.columns = df.columns.str.strip().str.lower()
            df.rename(columns={
                'date': 'Date', 
                'libellé': 'Libellé', 
                'montant(euros)': 'Montant(EUR)'
            }, inplace=True)

            for _, row in df.iterrows():
                cursor.execute('''
                    INSERT INTO transactions (date, libelle, montant) 
                    VALUES (?, ?, ?)
                ''', (row['Date'], row['Libellé'], row['Montant(EUR)']))

            conn.commit()
            conn.close()

            QMessageBox.information(self, 'Success', 'Data imported successfully!')

        except Exception as e:
            QMessageBox.critical(self, 'Error', str(e))


# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     ex = CsvToSqliteApp()
#     ex.show()
#     sys.exit(app.exec_())


df = pd.read_excel("C:/Users/joris/Downloads/2679266K0291720350549357.xlsx", engine='openpyxl')
print(df.columns)  # Affiche les noms de colonnes
