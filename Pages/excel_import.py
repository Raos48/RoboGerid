from tkinter import filedialog
from openpyxl import load_workbook

def import_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            workbook = load_workbook(filename=file_path, read_only=True)
            sheet = workbook.active
            return workbook, sheet, file_path
        except Exception as e:
            print(f"Erro ao importar o arquivo: {e}")
            return None, None, None
