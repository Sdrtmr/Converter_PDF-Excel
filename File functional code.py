import tkinter as tk
from tkinter import filedialog
import PyPDF2
import pandas as pd

def convert():
    pdf_path = filedialog.askopenfilename(filetypes=[('PDF Files', '*.pdf')])
    if pdf_path:
        excel_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
        if excel_path:
            pdf_file = open(pdf_path, 'rb')
            pdf_reader = PyPDF2.PdfFileReader(pdf_file)

            num_pages = pdf_reader.numPages
            rows = []

            for page in range(num_pages):
                page_obj = pdf_reader.getPage(page)
                content = page_obj.extract_text()

                # Разделение содержимого на строки
                lines = content.split('\n')
                rows.extend(lines)

            # Создание датафрейма из строк
            df = pd.DataFrame(rows)

            # Запись датафрейма в Excel-файл
            df.to_excel(excel_path, index=False)
            success_label.config(text='Конвертация выполнена успешно')

root = tk.Tk()
root.title('PDF to Excel Converter')

convert_btn = tk.Button(root, text='Выбрать PDF и сохранить в Excel', command=convert)
convert_btn.pack(pady=20)

success_label = tk.Label(root, text='')
success_label.pack()

root.mainloop()
