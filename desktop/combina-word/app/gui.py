from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .merge import detect_template_fields, load_records, merge_excel_into_one_docx, validate_template_fields


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Combina Word")
        self.geometry("780x390")
        self.minsize(700, 350)
        self.excel = tk.StringVar()
        self.template = tk.StringVar()
        self.copy_count = tk.StringVar(value="0")
        self.detected_count = tk.StringVar(value="Carga el Excel para detectar los registros.")
        self.detected_fields = tk.StringVar(value="Carga el DOCX para detectar sus categorías.")
        self.output = tk.StringVar()
        self.status = tk.StringVar(value="Selecciona los archivos para comenzar.")
        self.progress = tk.DoubleVar(value=0)
        self._build()

    def _build(self) -> None:
        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Combina Word", font=("TkDefaultFont", 18, "bold")).pack(anchor="w")
        ttk.Label(frame, text="Carga primero la plantilla Word y después Excel; la cantidad se detecta automáticamente.").pack(anchor="w", pady=(3, 20))
        self._file_row(frame, "Plantilla Word", self.template, self.pick_template, "Seleccionar plantilla")
        ttk.Label(frame, textvariable=self.detected_fields, foreground="#164e63").pack(anchor="w", pady=(0, 8))
        self._file_row(frame, "Archivo Excel", self.excel, self.pick_excel, "Seleccionar Excel")
        ttk.Label(frame, text="Cantidad de copias/páginas detectadas (puedes reducirla):").pack(anchor="w")
        ttk.Entry(frame, textvariable=self.copy_count, width=20).pack(anchor="w", pady=(3, 2))
        ttk.Label(frame, textvariable=self.detected_count, foreground="#164e63").pack(anchor="w", pady=(0, 12))
        self._file_row(frame, "Documento de salida", self.output, self.pick_output, "Guardar como")
        self.generate_button = ttk.Button(frame, text="Generar documento final", command=self.start)
        self.generate_button.pack(anchor="w", pady=(15, 8))
        ttk.Progressbar(frame, variable=self.progress, maximum=100).pack(fill="x")
        ttk.Label(frame, textvariable=self.status, foreground="#164e63").pack(anchor="w", pady=(8, 0))

    @staticmethod
    def _file_row(parent, label, variable, command, button_text):
        ttk.Label(parent, text=label).pack(anchor="w")
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(3, 12))
        ttk.Entry(row, textvariable=variable).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text=button_text, command=command).pack(side="left", padx=(8, 0))

    def pick_excel(self):
        if not self.template.get().strip():
            messagebox.showwarning("Orden requerido", "Primero carga la plantilla Word (.docx).")
            return
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls"), ("Todos", "*.*")])
        if path:
            try:
                count = len(load_records(path))
            except Exception as exc:
                messagebox.showerror("Excel inválido", f"No se pudo leer el Excel seleccionado:\n\n{exc}")
                return
            self.excel.set(path)
            self.copy_count.set(str(count))
            self.detected_count.set(f"Excel detectado: {count} registro(s); se generará una copia por registro.")
            matched, missing, _unused = validate_template_fields(self.template.get(), path)
            if missing:
                self.detected_fields.set("Faltan en Excel: " + ", ".join(missing))
            else:
                self.detected_fields.set("Categorías coincidentes: " + (", ".join(matched) if matched else "ninguna detectada"))

    def pick_template(self):
        path = filedialog.askopenfilename(filetypes=[("Word", "*.docx"), ("Todos", "*.*")])
        if path:
            try:
                fields = detect_template_fields(path)
            except Exception as exc:
                messagebox.showerror("Word inválido", f"No se pudo detectar el documento Word:\n\n{exc}")
                return
            self.template.set(path)
            self.detected_fields.set("Categorías detectadas en Word: " + (", ".join(fields) if fields else "ninguna"))

    def pick_output(self):
        path = filedialog.asksaveasfilename(
            title="Guardar documento combinado",
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
        )
        if path:
            self.output.set(path)

    def start(self):
        excel, template, output = map(str.strip, (self.excel.get(), self.template.get(), self.output.get()))
        raw_count = self.copy_count.get().strip()
        if not template:
            messagebox.showwarning("Falta la plantilla", "Primero carga la plantilla Word (.docx).")
            return
        if not raw_count.isdigit():
            messagebox.showwarning("Cantidad inválida", "Indica un número entero; usa 0 para todas las filas del Excel.")
            return
        copy_count = int(raw_count)
        if not excel or not template or not output:
            messagebox.showwarning("Datos incompletos", "Selecciona Excel, plantilla Word y documento de salida.")
            return
        if copy_count < 0:
            messagebox.showwarning("Cantidad inválida", "La cantidad de copias no puede ser negativa.")
            return
        if not Path(excel).is_file() or not Path(template).is_file():
            messagebox.showerror("Archivo no encontrado", "Verifica que Excel y la plantilla Word existan.")
            return
        _matched, missing, _unused = validate_template_fields(template, excel)
        if missing:
            messagebox.showerror("Categorías no encontradas", "El Excel no contiene estas categorías de la plantilla:\n\n" + ", ".join(missing))
            return
        if copy_count == 0:
            try:
                copy_count = len(load_records(excel))
                self.copy_count.set(str(copy_count))
            except Exception as exc:
                messagebox.showerror("Excel inválido", f"No se pudo detectar la cantidad de registros:\n\n{exc}")
                return
        if Path(output).resolve() in {Path(excel).resolve(), Path(template).resolve()}:
            messagebox.showerror("Salida inválida", "El documento de salida debe ser diferente a los archivos de entrada.")
            return
        self.generate_button.configure(state="disabled")
        self.progress.set(0)
        self.status.set("Generando documentos...")
        threading.Thread(target=self._worker, args=(excel, template, output, copy_count), daemon=True).start()

    def _worker(self, excel, template, output, record_limit):
        try:
            count = merge_excel_into_one_docx(excel, template, output, self._progress, record_limit)
            self.after(0, lambda: self._done(count, output))
        except Exception as exc:  # display friendly error in UI
            self.after(0, lambda: self._error(str(exc)))

    def _progress(self, current, total):
        self.after(0, lambda: self.progress.set(current * 100 / total))

    def _done(self, count, output):
        self.generate_button.configure(state="normal")
        self.status.set(f"Listo: {count} registro(s) combinados.")
        messagebox.showinfo("Proceso terminado", f"Se creó el documento final:\n\n{output}")

    def _error(self, message):
        self.generate_button.configure(state="normal")
        self.status.set("Ocurrió un error.")
        messagebox.showerror("No se pudo generar", message)


def main() -> None:
    App().mainloop()
