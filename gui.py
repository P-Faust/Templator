import customtkinter as tk
from CTkMessagebox import CTkMessagebox
from tkinter import filedialog
import excel_reader, word_writer, os
import customtkinter as tk
import os, traceback

def show_error(error_message):
    CTkMessagebox(title="Error", message=error_message)

if __name__ == "__main__":
    def select_excel_file():
        excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        excel_entry.delete(0, tk.END)
        excel_entry.insert(0, excel_file_path)

    def select_word_file():
        word_file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
        word_entry.delete(0, tk.END)
        word_entry.insert(0, word_file_path)

    def on_go():
        excel_file = excel_entry.get()
        word_file = word_entry.get()

        if os.path.exists(excel_file) and os.path.exists(word_file):
            try:
                perf = excel_reader.Perfomance(excel_file)
            except (RuntimeError, ValueError) as error:
                show_error(f"{error}")
                return
                

            roam = excel_reader.Roaming(excel_file)
            word_worker = word_writer.WordWriter(word_file, perf, roam)
            word_worker._write_document()
            CTkMessagebox(title="", message="Die Werte aus der Excel wurden in das Template übertragen.", icon="check")
        else:
            CTkMessagebox(title="Error", message="Datei nicht gefunden!", icon="cancel")

    root = tk.CTk()
    root.title("Templator!")
    root.eval('tk::PlaceWindow . center')


    # Erstelle die Eingabefelder für die Dateien
    excel_label = tk.CTkLabel(root, text="Excel Datei:")
    excel_label.grid(row=0, column=0, padx=10, pady=5, sticky="e")
    excel_entry = tk.CTkEntry(root, width=300)
    excel_entry.grid(row=0, column=1, padx=10, pady=5)
    excel_button = tk.CTkButton(root, text="Auswählen", command=select_excel_file)
    excel_button.grid(row=0, column=2, padx=10, pady=5)

    word_label = tk.CTkLabel(root, text="Word Datei:")
    word_label.grid(row=1, column=0, padx=10, pady=5, sticky="e")
    word_entry = tk.CTkEntry(root, width=300)
    word_entry.grid(row=1, column=1, padx=10, pady=5)
    word_button = tk.CTkButton(root, text="Auswählen", command=select_word_file)
    word_button.grid(row=1, column=2, padx=10, pady=5)

    # Go Button
    go_button = tk.CTkButton(root, text="Go", command=on_go)
    go_button.grid(row=2, column=1, padx=10, pady=10)

    root.mainloop()