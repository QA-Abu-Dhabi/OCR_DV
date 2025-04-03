# Для корректной компиляции кода в exe файл необходимо запустить следующую команду:
# pyinstaller --onefile --windowed --icon=favicon.ico --add-data "favicon.ico;." --add-data "poppler\bin;poppler\bin" --add-data "Tesseract-OCR;Tesseract-OCR" ocr.py
# Программа предназначена для распознавания дефектных ведомостей в формате pdf, извлечения определенных данных
# в соответствии с заданными регулярными выражениями и их сохранение в excel файле.

import sys
import re
import os
import pytesseract
import cv2
import numpy as np
import shutil
import tempfile
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QPushButton, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
from pdf2image import convert_from_path
from openpyxl import load_workbook

def get_poppler_path():
    """Получаем путь к каталогу, где находится EXE"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # Путь к временной директории EXE
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  # Обычный путь

    if getattr(sys, 'frozen', False):
        temp_dir = tempfile.mkdtemp()
        tesseract_src = os.path.join(sys._MEIPASS, "Tesseract-OCR")
        tesseract_dest = os.path.join(temp_dir, "Tesseract-OCR")
        shutil.copytree(tesseract_src, tesseract_dest, dirs_exist_ok=True)
        tesseract_path = os.path.join(tesseract_dest, "tesseract.exe")
    else:
        tesseract_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

    pytesseract.pytesseract.tesseract_cmd = tesseract_path

    return os.path.join(base_path, 'poppler', 'bin')  # Путь к poppler\bin рядом с EXE


def preprocess_image(image):
    """Обрабатывает изображение перед OCR, включая определение и исправление ориентации"""
    img = np.array(image)  # Конвертация в OpenCV
    img_gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)  # Преобразование в ч/б
    img_thresh = cv2.threshold(img_gray, 150, 255, cv2.THRESH_BINARY)[1]  # Бинаризация

    # Определяем ориентацию текста с помощью Tesseract
    osd = pytesseract.image_to_osd(img_thresh, output_type=pytesseract.Output.DICT)
    angle = osd["rotate"]  # Получаем угол поворота

    if angle != 0:
        img_thresh = cv2.rotate(img_thresh, {
            90: cv2.ROTATE_90_CLOCKWISE,
            180: cv2.ROTATE_180,
            270: cv2.ROTATE_90_COUNTERCLOCKWISE
        }[angle])

    return img_thresh


class PDFProcessor(QMainWindow):
    def __init__(self):
        super().__init__()
        
        if getattr(sys, 'frozen', False):  # Проверяем, запущен ли exe-файл
            base_path = sys._MEIPASS  # Временная папка PyInstaller
        else:
            base_path = os.path.dirname(__file__)  # Обычный путь для скрипта

        icon_path = os.path.join(base_path, "favicon.ico")
        self.setWindowIcon(QIcon(icon_path))
        
        self.setWindowTitle("Распознавание Дефектных ведомостей")
        self.setGeometry(300, 200, 600, 500)

        self.poppler_path = get_poppler_path()
        self.selected_folder = ""

        layout = QVBoxLayout()

        # Добавляем описание программы в QLabel
        self.label_description = QLabel(
            "Данная программа предназначена для распознавания дефектных\n"
            "ведомостей и сохранения ключевых параметров в файл Excel.\n\n"
            "Для распознавания необходимо выбрать папку с дефектными ведомостями в PDF формате."
        )
        self.label_description.setWordWrap(True)  # Разрешаем перенос строк
        self.label_description.setAlignment(Qt.AlignLeft)  # Выравниваем текст по левому краю
        layout.addWidget(self.label_description)  # Добавляем в макет
        
        # Кнопка выбора папки
        self.btn_select_folder = QPushButton("Выбрать папку с ДВ")
        self.btn_select_folder.clicked.connect(self.select_folder)
        self.btn_select_folder.setFixedSize(300, 70)
        layout.addWidget(self.btn_select_folder, alignment=Qt.AlignCenter)

        # Метка для отображения пути к выбранной папке
        self.label_folder_path = QLabel("Папка не выбрана", self)
        self.label_folder_path.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label_folder_path)

        # Кнопка распознавания
        self.btn_process = QPushButton("Распознать ДВ и сохранить в excel")
        self.btn_process.clicked.connect(self.process_pdfs)
        self.btn_process.setEnabled(False)  # Отключена, пока не выбрана папка
        self.btn_process.setFixedSize(300, 70)
        layout.addWidget(self.btn_process, alignment=Qt.AlignCenter)

        layout.setSpacing(40)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)
        
        layout.addStretch() # Прижимаем текст к нижнему краю
        self.label_author = QLabel("© 2025 Автор: Абубакиров Ильмир Иргалиевич", self)
        self.label_author.setAlignment(Qt.AlignRight)  # Выравнивание справа
        layout.addWidget(self.label_author)
        self.label_author.setStyleSheet("font-size: 12px; color: gray;")

    def select_folder(self):
        """Открывает диалоговое окно для выбора папки"""
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с ДВ", "")
        if folder:
            self.selected_folder = folder
            self.label_folder_path.setText(f"Выбрана папка:\n{folder}")
            self.btn_process.setEnabled(True)  # Включаем кнопку распознавания

    def process_pdfs(self):
        """Обрабатывает все PDF в выбранной папке, извлекает числа и сохраняет в Excel"""
        if not self.selected_folder:
            QMessageBox.warning(self, "Ошибка", "Выберите папку перед началом обработки!")
            return

        pdf_files = [f for f in os.listdir(self.selected_folder) if f.lower().endswith('.pdf')]
        if not pdf_files:
            QMessageBox.warning(self, "Ошибка", "В выбранной папке нет PDF-файлов!")
            return

        results = []  # Список для хранения результатов

        for pdf_file in pdf_files:
            pdf_path = os.path.join(self.selected_folder, pdf_file)
            print(f"Обрабатывается файл: {pdf_path}")

            try:
                images = convert_from_path(pdf_path, poppler_path=self.poppler_path)
                full_text = ""

                for image in images:
                    processed_image = preprocess_image(image)
                    text = pytesseract.image_to_string(processed_image, lang='rus+eng', config='--psm 6')
                    full_text += text + "\n"

                # Очистить текст от лишних символов
                full_text = full_text.replace("\n", " ").replace("\t", " ")

                # Ищем 12-значное число
                match0 = re.search(r"\b\d{12}\b", full_text)
                ZakazTOPO = match0.group(0) if match0 else "Не найдено"
                
                # Поиск ЛПУМГ
                match1 = re.search(r"\b([А-Яа-яЁёA-Za-z-]+)\sЛПУМГ\b", full_text)
                found_lpumg = match1.group(1) if match1 else "Не найдено"

                # Поиск Инвентарный номер
                match2 = re.search(r"Инвентарный №:\s*(\S+)", full_text)
                found_inventory = match2.group(1) if match2 else "Не найдено"

                # Поиск Объект ремонта
                match3 = re.search(r"На капитальный ремонт объекта\s*[-—–]?\s*(.*?)\s*\(?\s*\{?\s*название объекта в соответствии", full_text, re.IGNORECASE | re.DOTALL)
                
                if not match3:  # если первое не нашло совпадение, проверяем второе
                    match3 = re.search(r"На капитальный ремонт объекта\s*[-—–]?\s*(.*?)\s*\(?\s*\{?\s*содержание выполняемых работ", full_text, re.IGNORECASE | re.DOTALL)
                
                found_repair_object = match3.group(1) if match3 else "Не найдено"
                
                results.append([pdf_file, ZakazTOPO, found_lpumg, found_inventory, found_repair_object])

            except Exception as e:
                results.append([pdf_file, f"Ошибка: {str(e)}"])

        # Сохранение результатов в Excel
        output_file = os.path.join(self.selected_folder, "результаты.xlsx")
        df = pd.DataFrame(results, columns=["Файл", "заказ ТОРО", "Филиал", "Инв. №", "Наименование объекта"])
        df.to_excel(output_file, index=False, engine='openpyxl')

        # Открываем созданный файл для изменения ширины столбцов
        wb = load_workbook(output_file)
        ws = wb.active

        # Настраиваем ширину столбцов
        column_widths = {
            "A": 20,  # "Файл"
            "B": 15,  # "заказ ТОРО"
            "C": 19,  # "Филиал"
            "D": 16,  # "Инв. №"
            "E": 60   # "Наименование объекта"
        }

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width

        # Сохраняем изменения
        wb.save(output_file)
        wb.close()
        
        QMessageBox.information(self, "Готово", f"Обработка завершена!\nРезультаты сохранены в:\n{output_file}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFProcessor()
    window.show()
    sys.exit(app.exec_())
