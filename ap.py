import re
import pandas as pd
import fitz
from tkinter import *
from tkinter import filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD


# ---------------- PDF PARSER ----------------

def parse_pdf(pdf_path):

    rows = []

    doc = fitz.open(pdf_path)

    for page in doc:
        text = page.get_text("text")

        pos = re.search(r"Позиция:\s*(\S+)", text)
        desc = re.search(r"Описание:\s*(.+)", text)
        system = re.search(r"Система:\s*(.+)", text)
        color = re.search(r"Цвет:\s*(\S+)", text)
        frame = re.search(r"Рама:\s*([\d\s]+)x([\d\s]+)", text)
        qty = re.search(r"Кол-во:\s*(\d+)", text)

        if not pos:
            continue

        name = pos.group(1)

        description = desc.group(1).strip().lower() if desc else ""

        system_name = ""
        if system:
            s = system.group(1)
            s = s.replace("AluminTechno", "Alutech")
            m = re.search(r"(Alutech\s+[A-Z]*\d+)", s)
            if m:
                system_name = m.group(1)

        color_val = color.group(1) if color else ""

        width = int(frame.group(1).replace(" ", "")) if frame else None
        height = int(frame.group(2).replace(" ", "")) if frame else None

        quantity = int(qty.group(1)) if qty else 1

        rows.append([
            name,
            description,
            system_name,
            "",
            color_val,
            "",
            width,
            height,
            quantity
        ])

    return rows


# ---------------- GUI ----------------

class App:

    def __init__(self, root):

        self.root = root
        self.root.title("AluPro PDF to Excel")
        self.root.geometry("660x240")

        self.pdf_path = ""
        self.save_path = ""

        # PDF
        frame_pdf = Frame(root)
        frame_pdf.pack(pady=8, padx=10, fill="x")

        Label(frame_pdf, text="Файл PDF:", width=10, anchor="w").pack(side="left")

        self.pdf_entry = Entry(frame_pdf)
        self.pdf_entry.pack(side="left", fill="x", expand=True, padx=6)

        Button(frame_pdf, text="Выбрать", command=self.choose_pdf).pack(side="left")

        # SAVE
        frame_save = Frame(root)
        frame_save.pack(pady=8, padx=10, fill="x")

        Label(frame_save, text="Excel:", width=10, anchor="w").pack(side="left")

        self.save_entry = Entry(frame_save)
        self.save_entry.pack(side="left", fill="x", expand=True, padx=6)

        Button(frame_save, text="Куда сохранить", command=self.choose_save).pack(side="left")

        # DROP ZONE
        self.drop_label = Label(root, text="Перетащи PDF сюда",
                                bg="#dddddd", height=5, relief="ridge")
        self.drop_label.pack(fill="x", padx=10, pady=10)

        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.drop_pdf)

        # PROCESS
        Button(root, text="ОБРАБОТАТЬ",
               height=1, font=("Arial", 11, "bold"),
               command=self.process).pack(pady=6)

    # --------- actions ---------

    def choose_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.pdf_path = path
            self.pdf_entry.delete(0, END)
            self.pdf_entry.insert(0, path)

    def choose_save(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx")])
        if path:
            self.save_path = path
            self.save_entry.delete(0, END)
            self.save_entry.insert(0, path)

    def drop_pdf(self, event):

        path = event.data.strip("{}")

        if not path.lower().endswith(".pdf"):
            messagebox.showerror("Ошибка", "Можно только PDF")
            return

        self.pdf_path = path
        self.pdf_entry.delete(0, END)
        self.pdf_entry.insert(0, path)

        self.drop_label.config(text="PDF загружен ✓", bg="#ccffcc")

    def process(self):

        if not self.pdf_path:
            messagebox.showerror("Ошибка", "Не выбран PDF")
            return

        if not self.save_path:
            messagebox.showerror("Ошибка", "Не указан путь сохранения")
            return

        data = parse_pdf(self.pdf_path)
        df = pd.DataFrame(data)

        df.to_excel(self.save_path, index=False, header=False)

        messagebox.showinfo("Готово", "Excel сохранен")


# ---------------- START ----------------

root = TkinterDnD.Tk()
root.resizable(False, False)  # запрет изменения размера окна
app = App(root)
root.mainloop()
