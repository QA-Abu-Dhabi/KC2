# Для корректной компиляции кода в exe файл необходимо запустить следующую команду:
# pyinstaller --onefile --windowed --icon=favicon.ico --add-data "favicon.ico;." KC2.py

import os
import re
import sys
from PyQt5.QtWidgets import (QApplication, QMainWindow, QLabel, QVBoxLayout, 
                             QWidget, QPushButton, QFileDialog, QMessageBox, 
                             QProgressBar, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
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
        
        result = {
            'Название файла': os.path.basename(pdf_path),
            'Сумма по КС2': None,
            'Номер документа': None,
            'Код АСУ ТОиР': None,
            'Код X/Y/Z': None,
            'Ошибка': None
        }

        try:
            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                print(f"Всего страниц в документе: {total_pages}")
                
                # 1. Поиск "Всего по акту" и суммы (со всех страниц с конца)
                for page_num in range(total_pages-1, -1, -1):
                    current_page = pdf.pages[page_num]
                    print(f"\nПоиск суммы: анализируем страницу {page_num+1}/{total_pages}")
                    
                    tables = current_page.extract_tables({"vertical_strategy": "lines", "horizontal_strategy": "text"})
                    
                    for table_num, table in enumerate(tables or [], 1):
                        for row_num, row in enumerate(table or []):
                            for col_num, cell in enumerate(row or []):
                                if cell and re.search(r'всего по акту', str(cell), re.IGNORECASE):
                                    print(f"Найдено 'Всего по акту' в таблице {table_num}, строка {row_num}, столбец {col_num}")
                                    
                                    # Поиск числа справа
                                    for right_col in range(col_num+1, len(row)):
                                        try:
                                            value = float(str(row[right_col]).replace(',', '.').replace(' ', ''))
                                            result.update({
                                                'Сумма по КС2': value,
                                            })
                                            print(f"Найдена сумма: {value}")
                                            break
                                        except (ValueError, TypeError):
                                            continue
                                    else:
                                        result['Ошибка'] = "Не найдена сумма рядом с 'Всего по акту'"
                                    break
                            else:
                                continue
                            break
                        else:
                            continue
                        break
                    else:
                        continue
                    break
                
                # 2. Поиск дополнительных данных на первой странице (page_num=0)
                if total_pages > 0:
                    first_page = pdf.pages[0]
                    print("\nАнализируем первую страницу для поиска дополнительных данных")
                    
                    # 2.1. Поиск "Номер документа" и текста под ним
                    first_page_tables = first_page.extract_tables()
                    for table in first_page_tables or []:
                        for row_idx, row in enumerate(table or []):
                            for col_idx, cell in enumerate(row or []):
                                if cell and re.search(r'документа', str(cell), re.IGNORECASE):
                                    print(f"Найдена строка с 'Номер документа' в столбце {col_idx}")
                                    
                                    # Ищем первое непустое значение ниже в том же столбце
                                    for next_row_idx in range(row_idx+1, len(table)):
                                        next_row = table[next_row_idx]
                                        if col_idx < len(next_row) and next_row[col_idx] and str(next_row[col_idx]).strip():
                                            result['Номер документа'] = str(next_row[col_idx]).strip()
                                            print(f"Найден номер документа: {result['Номер документа']}")
                                            break
                                    
                                    if not result['Номер документа']:
                                        result['Ошибка'] = "Не найден номер документа под заголовком"
                                    break
                            else:
                                continue
                            break
                        else:
                            continue
                        break
                    
                    # 2.2. Поиск кодов в тексте первой страницы
                    first_page_text = first_page.extract_text() or ""
                    if first_page_text:
                        # Код АСУ ТОиР (14 цифр, начинается с 2)
                        asu_match = re.search(r'(?<!\d)2\d{13}(?!\d)', first_page_text)
                        if asu_match:
                            result['Код АСУ ТОиР'] = asu_match.group(0)
                            print(f"Найден Код АСУ ТОиР: {result['Код АСУ ТОиР']}")
                        
                        # Код X/Y/Z (формат */*/*)
                        xyz_match = re.search(r'\b(\d{4,})[/-]([\d.]+)[/-](\d{4,})\b', first_page_text)
                        if xyz_match:
                            result['Код X/Y/Z'] = f"{xyz_match.group(1)}/{xyz_match.group(2)}/{xyz_match.group(3)}"
                            print(f"Найден Код X/Y/Z: {result['Код X/Y/Z']}")
        
        except Exception as e:
            error_msg = f"Ошибка обработки файла: {str(e)}"
            print(error_msg)
            result['Ошибка'] = error_msg
        
        print("\nИтоговые результаты:")
        for key, value in result.items():
            print(f"{key}: {value}")
        
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
        self.setWindowTitle("Анализ форм КС-2")
        self.setGeometry(100, 100, 700, 500)
        
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
        
        # Добавляем описание программы в QLabel
        self.label_description = QLabel(
            "Программа предназначена для автоматического анализа отчетных форм КС-2,\n"
            "поиска финансовых данных в таблицах и сохранения результатов в файл Excel\n\n"
            "Для начала работы необходимо выбрать папку с файлами КС-2 в PDF формате.\n\n\n"
        )
        self.label_description.setWordWrap(True)  # Разрешаем перенос строк
        self.label_description.setAlignment(Qt.AlignLeft)  # Выравниваем текст по левому краю
        layout.addWidget(self.label_description)  # Добавляем в макет
        
        self.status_label = QLabel("Выберите папку для анализа КС-2")
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)
        
        self.detailed_status = QLabel("")
        self.detailed_status.setAlignment(Qt.AlignCenter)
        self.detailed_status.setWordWrap(True)  # Разрешаем перенос слов
        self.detailed_status.setMinimumWidth(400)  # Устанавливаем минимальную ширину
        layout.addWidget(self.detailed_status)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)  # Скрываем прогресс-бар по умолчанию
        layout.addWidget(self.progress_bar)
        
        self.select_dir_btn = QPushButton("Выбрать папку")
        self.select_dir_btn.clicked.connect(self.select_directory)
        self.select_dir_btn.setFixedSize(400, 70)
        layout.addWidget(self.select_dir_btn, alignment=Qt.AlignCenter)
        
        self.process_btn = QPushButton("Обработать файлы")
        self.process_btn.clicked.connect(self.process_files)
        self.process_btn.setEnabled(False)
        self.process_btn.setFixedSize(400, 70)
        layout.addWidget(self.process_btn, alignment=Qt.AlignCenter)
        
        # Добавляем чекбокс для открытия файла после обработки
        self.open_file_checkbox = QCheckBox("Открыть файл после завершения обработки")
        self.open_file_checkbox.setChecked(True)  # Устанавливаем галочку по умолчанию
        layout.addWidget(self.open_file_checkbox)
        
        # Добавляем растягивающийся элемент, чтобы прижать элементы к верху
        layout.addStretch()
        
        layout.addStretch() # Прижимаем текст к нижнему краю
        self.label_author = QLabel("© 2025 Автор: Абубакиров Ильмир Иргалиевич", self)
        self.label_author.setAlignment(Qt.AlignRight)  # Выравнивание справа
        layout.addWidget(self.label_author)
        self.label_author.setStyleSheet("font-size: 12px; color: gray;")
        
        if getattr(sys, 'frozen', False):  # Проверяем, запущен ли exe-файл
            base_path = sys._MEIPASS  # Временная папка PyInstaller
        else:
            base_path = os.path.dirname(__file__)  # Обычный путь для скрипта
            
        icon_path = os.path.join(base_path, "favicon.ico")
        self.setWindowIcon(QIcon(icon_path))
    
    def select_directory(self):
        """Выбор директории для поиска PDF файлов"""
        directory = QFileDialog.getExistingDirectory(self, "Выберите директорию")
        if directory:
            self.directory = directory
            self.status_label.setText(f"Выбрана папка: {directory}")
            self.process_btn.setEnabled(True)
    
    def process_files(self):
        """Обработка файлов в выбранной директории"""
        if hasattr(self, 'directory'):
            self.status_label.setText("Поиск PDF файлов...")
            self.detailed_status.setText("")
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)  # Показываем прогресс-бар перед началом обработки
            
            QApplication.processEvents()
            
            self.extractor.directory = self.directory
            self.extractor.start()
            
            self.process_btn.setEnabled(False)
            self.select_dir_btn.setEnabled(False)
    
    def process_finished(self, results):
        """Завершение обработки файлов"""
        self.process_btn.setEnabled(True)
        self.select_dir_btn.setEnabled(True)
        self.progress_bar.setVisible(False)  # Скрываем прогресс-бар после завершения
        
        if results:
            self.status_label.setText(f"Обработано {len(results)} файлов. Результаты сохранены в Result.xlsx")
            # Если чекбокс отмечен, открываем файл
            if self.open_file_checkbox.isChecked():
                output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Result.xlsx')
                if os.path.exists(output_path):
                    try:
                        os.startfile(output_path)  # Для Windows
                    except:
                        import subprocess
                        try:
                            # Попробуем открыть файл для других ОС
                            opener = "open" if sys.platform == "darwin" else "xdg-open"
                            subprocess.call([opener, output_path])
                        except:
                            QMessageBox.information(self, "Файл сохранен", 
                                                  f"Файл сохранен по пути:\n{output_path}")
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