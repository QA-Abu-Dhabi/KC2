import os
import re
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
import pandas as pd
import pdfplumber


class PDFTableExtractor:
    def __init__(self):
        self.found_files = []
        self.results = []

    def find_pdf_files(self, directory):
        """Поиск PDF файлов с указанными паттернами в названии"""
        patterns = [r'KC2', r'KC-2', r'KC_2']
        self.found_files = []
        
        for root, _, files in os.walk(directory):
            for file in files:
                if file.lower().endswith('.pdf'):
                    for pattern in patterns:
                        if re.search(pattern, file, re.IGNORECASE):
                            self.found_files.append(os.path.join(root, file))
                            break
        
        return self.found_files

    def extract_value_from_table(self, pdf_path):
        """Извлечение значения из таблицы в PDF файле"""
        with pdfplumber.open(pdf_path) as pdf:
            # Проверяем страницы с конца
            for page in reversed(pdf.pages):
                tables = page.extract_tables()
                if tables:
                    # Берем последнюю таблицу на странице
                    last_table = tables[-1]
                    if last_table and len(last_table) > 0:
                        # Берем последнюю строку таблицы
                        last_row = last_table[-1]
                        if last_row and len(last_row) > 0:
                            # Берем последнее значение в строке
                            last_value = last_row[-1]
                            try:
                                # Пробуем преобразовать в число
                                return float(last_value)
                            except (ValueError, TypeError):
                                # Если не получается, возвращаем как есть
                                return last_value
        return None

    def process_files(self, directory):
        """Обработка всех найденных файлов"""
        self.find_pdf_files(directory)
        self.results = []
        
        for file in self.found_files:
            value = self.extract_value_from_table(file)
            if value is not None:
                self.results.append({
                    'File': os.path.basename(file),
                    'Path': file,
                    'Value': value
                })
        
        return self.results

    def save_to_excel(self, output_path):
        """Сохранение результатов в Excel файл"""
        if not self.results:
            return False
        
        df = pd.DataFrame(self.results)
        try:
            df.to_excel(output_path, index=False)
            return True
        except Exception as e:
            print(f"Error saving Excel file: {e}")
            return False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Table Extractor")
        self.setGeometry(100, 100, 400, 200)
        
        self.extractor = PDFTableExtractor()
        self.init_ui()
    
    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        self.status_label = QLabel("Выберите директорию для поиска PDF файлов")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.select_dir_btn = QPushButton("Выбрать директорию")
        self.select_dir_btn.clicked.connect(self.select_directory)
        layout.addWidget(self.select_dir_btn)
        
        self.process_btn = QPushButton("Обработать файлы")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)
        
        self.save_btn = QPushButton("Сохранить в Excel")
        self.save_btn.clicked.connect(self.save_results)
        self.save_btn.setEnabled(False)
        layout.addWidget(self.save_btn)
    
    def select_directory(self):
        """Выбор директории для поиска PDF файлов"""
        directory = QFileDialog.getExistingDirectory(self, "Выберите директорию")
        if directory:
            self.directory = directory
            self.status_label.setText(f"Выбрана директория: {directory}")
            self.process_btn.setEnabled(True)
    
    def process_files(self):
        """Обработка файлов в выбранной директории"""
        if hasattr(self, 'directory'):
            self.status_label.setText("Обработка файлов...")
            QApplication.processEvents()  # Обновляем интерфейс
            
            results = self.extractor.process_files(self.directory)
            
            if results:
                self.status_label.setText(f"Найдено {len(results)} файлов с таблицами")
                self.save_btn.setEnabled(True)
            else:
                self.status_label.setText("Не найдено файлов с таблицами")
                self.save_btn.setEnabled(False)
    
    def save_results(self):
        """Сохранение результатов в Excel файл"""
        if not self.extractor.results:
            QMessageBox.warning(self, "Ошибка", "Нет данных для сохранения")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как Excel", "", "Excel Files (*.xlsx)"
        )
        
        if file_path:
            if not file_path.endswith('.xlsx'):
                file_path += '.xlsx'
            
            success = self.extractor.save_to_excel(file_path)
            if success:
                QMessageBox.information(self, "Успех", "Файл успешно сохранен")
                self.status_label.setText(f"Результаты сохранены в: {file_path}")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось сохранить файл")


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()