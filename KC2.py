import os
import re
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, 
                             QWidget, QPushButton, QFileDialog, QMessageBox, 
                             QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd
import pdfplumber
from openpyxl.styles import numbers


class PDFTableExtractor(QThread):
    progress_updated = pyqtSignal(int)
    processing_message = pyqtSignal(str)
    finished_processing = pyqtSignal(list)

    def __init__(self):
        super().__init__()
        self.found_files = []
        self.results = []
        self.directory = ""
        self.running = False
        self.total_keywords = ['всего', 'итого']

    def find_pdf_files(self):
        """Поиск PDF файлов с указанными паттернами в названии"""
        patterns = [r'KC2', r'KC-2', r'KC_2', r'КС2', r'КС-2', r'КС_2']
        self.found_files = []
        
        for root, _, files in os.walk(self.directory):
            for file in files:
                if file.lower().endswith('.pdf'):
                    if any(re.search(pattern, file, re.IGNORECASE) for pattern in patterns):
                        self.found_files.append(os.path.join(root, file))
        
        return self.found_files

    def extract_data_from_pdf(self, pdf_path):
        """Извлечение всех требуемых данных из PDF файла"""
        result = {
            'Название файла': os.path.basename(pdf_path),
            'Сумма по КС2': None,
            'Код АСУ ТОиР': None,
            'Код X/Y/Z': None,
            'Номер документа': None
        }

        try:
            with pdfplumber.open(pdf_path) as pdf:
                # Сначала ищем все текстовые данные
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        # Поиск 14-значного числа, начинающегося с 2 (Код АСУ ТОиР)
                        asu_match = re.search(r'(?<!\d)2\d{13}(?!\d)', text)
                        if asu_match:
                            result['Код АСУ ТОиР'] = asu_match.group(0)

                        # Поиск шаблона X/Y/Z с длинами частей >=4
                        xyz_match = re.search(r'\b(\d{4,})/(\d+)/(\d{4,})\b', text)
                        if xyz_match:
                            result['Код X/Y/Z'] = f"{xyz_match.group(1)}/{xyz_match.group(2)}/{xyz_match.group(3)}"

                # Затем ищем табличные данные
                for page_num in range(len(pdf.pages)-1, -1, -1):
                    page = pdf.pages[page_num]
                    self.processing_message.emit(f"Обработка {os.path.basename(pdf_path)}, страница {page_num+1}")
                    
                    # Поиск таблицы с "Отчетный период"
                    tables = page.extract_tables({
                        "vertical_strategy": "lines", 
                        "horizontal_strategy": "text"
                    })
                    
                    if tables:
                        for table in tables:
                            if not table:
                                continue
                            
                            # Поиск "Отчетный период" в таблице
                            for row in table:
                                for cell in row:
                                    if cell and 'отчетный период' in str(cell).lower():
                                        # Берем первое значение из последней строки таблицы
                                        if table and table[-1] and len(table[-1]) > 0:
                                            result['Номер документа'] = table[-1][0]
                                        break
                            
                            # Поиск таблицы с "Всего" или "Итого"
                            found_keyword = False
                            for row in table:
                                for cell in row:
                                    if cell and any(keyword in str(cell).lower() for keyword in self.total_keywords):
                                        found_keyword = True
                                        break
                                if found_keyword:
                                    break
                            
                            if found_keyword:
                                # Ищем число в последней строке последнего столбца
                                last_row = table[-1]
                                if last_row:
                                    for cell in reversed(last_row):
                                        if cell:
                                            try:
                                                value = float(str(cell).replace(',', '.').replace(' ', ''))
                                                result['Сумма по КС2'] = value
                                                break
                                            except (ValueError, TypeError):
                                                continue
                    
                    # Если все данные найдены, можно прекращать поиск
                    if all(v is not None for v in result.values() if v != 'Название файла'):
                        break
        
        except Exception as e:
            self.processing_message.emit(f"Ошибка при обработке {pdf_path}: {str(e)}")
        
        return result

    def run(self):
        """Обработка всех найденных файлов"""
        self.running = True
        self.results = []
        found_files = self.find_pdf_files()
        
        if not found_files:
            self.finished_processing.emit([])
            return
        
        for i, file in enumerate(found_files):
            if not self.running:
                break
                
            self.progress_updated.emit(int((i + 1) / len(found_files) * 100))
            file_data = self.extract_data_from_pdf(file)
            
            if any(v is not None for k, v in file_data.items() if k != 'Название файла'):
                self.results.append(file_data)
        
        self.save_results()
        self.finished_processing.emit(self.results)
    
    def save_results(self):
        """Автоматическое сохранение результатов в Excel файл"""
        if not self.results:
            return False
        
        try:
            df = pd.DataFrame(self.results)
            
            # Устанавливаем порядок столбцов
            column_order = [
                'Название файла',
                'Сумма по КС2',
                'Код АСУ ТОиР',
                'Код X/Y/Z',
                'Номер документа'
            ]
            df = df[column_order]
            
            # Формируем путь к файлу в директории с программой
            output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Result.xlsx')
            
            # Создаем Excel writer с настройками
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                
                # Получаем доступ к листу для настройки
                worksheet = writer.sheets['Sheet1']
                
                # Устанавливаем ширину столбцов
                worksheet.column_dimensions['A'].width = 65   # Название файла
                worksheet.column_dimensions['B'].width = 15   # Сумма по КС2
                worksheet.column_dimensions['C'].width = 17   # Код АСУ ТОиР
                worksheet.column_dimensions['D'].width = 15   # Код X/Y/Z
                worksheet.column_dimensions['E'].width = 16   # Номер документа
                
                # Устанавливаем числовой формат для столбца "Сумма по КС2"
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=2)
                    cell.number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED2
            
            return True
        except Exception as e:
            self.processing_message.emit(f"Ошибка при сохранении: {str(e)}")
            return False
    
    def stop(self):
        self.running = False


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Table Extractor")
        self.setGeometry(100, 100, 700, 400)
        
        self.extractor = PDFTableExtractor()
        self.init_ui()
        
        # Подключаем сигналы
        self.extractor.progress_updated.connect(self.update_progress)
        self.extractor.processing_message.connect(self.update_status)
        self.extractor.finished_processing.connect(self.process_finished)
    
    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        layout = QVBoxLayout()
        central_widget.setLayout(layout)
        
        self.status_label = QLabel("Выберите директорию для поиска PDF файлов")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.detailed_status = QLabel("")
        self.detailed_status.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.detailed_status)
        
        self.progress_bar = QProgressBar()
        layout.addWidget(self.progress_bar)
        
        self.select_dir_btn = QPushButton("Выбрать директорию")
        self.select_dir_btn.clicked.connect(self.select_directory)
        layout.addWidget(self.select_dir_btn)
        
        self.process_btn = QPushButton("Обработать файлы")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)
    
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
            self.status_label.setText("Поиск PDF файлов...")
            self.detailed_status.setText("")
            self.progress_bar.setValue(0)
            QApplication.processEvents()
            
            self.extractor.directory = self.directory
            self.extractor.start()
            
            self.process_btn.setEnabled(False)
            self.select_dir_btn.setEnabled(False)
    
    def process_finished(self, results):
        """Завершение обработки файлов"""
        self.process_btn.setEnabled(True)
        self.select_dir_btn.setEnabled(True)
        
        if results:
            self.status_label.setText(f"Обработано {len(results)} файлов. Результаты сохранены в Result.xlsx")
        else:
            self.status_label.setText("Не найдено файлов с требуемыми данными")
    
    def update_progress(self, value):
        """Обновление прогрессбара"""
        self.progress_bar.setValue(value)
    
    def update_status(self, message):
        """Обновление детального статуса"""
        self.detailed_status.setText(message)
        QApplication.processEvents()
    
    def closeEvent(self, event):
        """Обработка закрытия окна"""
        if self.extractor.isRunning():
            self.extractor.stop()
            self.extractor.wait()
        event.accept()


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec_()