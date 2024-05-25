import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os

def select_docx_file():
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo DOCX",
        filetypes=[("Documentos de Word", "*.docx")])
    if file_path:
        entry_docx_path.delete(0, tk.END)
        entry_docx_path.insert(0, file_path)

def select_output_directory():
    directory = filedialog.askdirectory(title="Seleccionar directorio de salida")
    if directory:
        entry_output_directory.delete(0, tk.END)
        entry_output_directory.insert(0, directory)

def convert_to_pdf():
    docx_path = entry_docx_path.get()
    output_directory = entry_output_directory.get()
    if not docx_path or not output_directory:
        messagebox.showwarning("Advertencia", "Por favor, selecciona un archivo DOCX y un directorio de salida.")
        return

    try:
        output_path = os.path.join(output_directory, os.path.basename(docx_path).replace(".docx", ".pdf"))
        convert(docx_path, output_path)
        messagebox.showinfo("Éxito", f"Archivo convertido y guardado en: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo convertir el archivo: {e}")

app = tk.Tk()
app.title("Convertidor de DOCX a PDF")

tk.Label(app, text="Archivo DOCX:").grid(row=0, column=0, padx=10, pady=10)
entry_docx_path = tk.Entry(app, width=50)
entry_docx_path.grid(row=0, column=1, padx=10, pady=10)
tk.Button(app, text="Seleccionar", command=select_docx_file).grid(row=0, column=2, padx=10, pady=10)

tk.Label(app, text="Directorio de salida:").grid(row=1, column=0, padx=10, pady=10)
entry_output_directory = tk.Entry(app, width=50)
entry_output_directory.grid(row=1, column=1, padx=10, pady=10)
tk.Button(app, text="Seleccionar", command=select_output_directory).grid(row=1, column=2, padx=10, pady=10)

tk.Button(app, text="Convertir a PDF", command=convert_to_pdf).grid(row=2, column=0, columnspan=3, pady=20)

app.mainloop()
