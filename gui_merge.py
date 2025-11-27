"""
Interfaz gráfica para merge de archivos F42
Principios: SOLID, Clean Code, KISS, DRY
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import tempfile
import shutil
import sys


class F42MergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Merge F42 - Múltiples Sedes")
        self.excel_files = []
        self.output_path = None
        # Estilos base — similar a la webapp (fondo claro, fuente legible)
        self.root.configure(bg='#f8f9fa')
        default_font = ('Segoe UI', 10)
        self.root.option_add('*Font', default_font)
        self._build_ui()

    def _build_ui(self):
        # Minimal layout per user request
        self.root.configure(bg='#f0f0f0')
        frame = tk.Frame(self.root, bg='#ffffff', padx=12, pady=12)
        frame.pack(padx=20, pady=20, fill='both', expand=True)

        # Button 1: choose files (white bg, black text)
        self.choose_btn = tk.Button(
            frame, 
            text='Adjuntar archivos F42 Excel', 
            command=self.add_excel_files, 
            bg='white', 
            fg='black'
        )
        self.choose_btn.pack(anchor='w', pady=(0, 8))

        # List of attached files
        self.excel_listbox = tk.Listbox(frame, width=80, height=6)
        self.excel_listbox.pack(anchor='w', pady=(0, 8))

        # Button 2: generate (match choose button style: white bg, black text)
        self.generate_btn = tk.Button(
            frame, 
            text='Generar F42 Merged', 
            command=self.run_merge, 
            bg='white', 
            fg='black'
        )
        self.generate_btn.pack(anchor='w', pady=(6, 8))

        # Success message (initially empty)
        self.status_label = tk.Label(frame, text='', bg='#ffffff', fg='black')
        self.status_label.pack(anchor='w', pady=(4, 8))

        # Button 3: download (match choose button style: white bg, black text)
        self.open_btn = tk.Button(
            frame, 
            text='Descargar F42 Merged', 
            command=self.open_output, 
            state=tk.NORMAL, 
            bg='white', 
            fg='black'
        )
        self.open_btn.pack(anchor='w', pady=(6, 0))

    def add_excel_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Archivos Excel", "*.xlsx *.xls")]
        )
        if files:
            # Acumular archivos sin duplicados
            new_files = [f for f in files if f not in self.excel_files]
            self.excel_files.extend(new_files)
            self.excel_listbox.delete(0, tk.END)
            for f in self.excel_files:
                self.excel_listbox.insert(tk.END, os.path.basename(f))
            # Mostrar tooltip con ruta completa
            self._add_tooltip(self.excel_listbox, self.excel_files)

    def _add_tooltip(self, widget, items):
        # simple tooltip: al pasar por encima muestra ruta completa en title
        def on_enter(event):
            if items:
                widget.master.title(f"Archivos seleccionados: {len(items)}")
        def on_leave(event):
            widget.master.title("Merge F42 - Múltiples Sedes")
        widget.bind('<Enter>', on_enter)
        widget.bind('<Leave>', on_leave)

    def read_f42_file(self, file_path):
        """Lee un archivo F42 Excel usando header=1 (row 1 es el header real)"""
        try:
            df = pd.read_excel(file_path, engine='openpyxl', header=1)
            # Drop any 'Unnamed' columns
            df = df.loc[:, ~df.columns.str.contains('^Unnamed', na=False)]
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except Exception as e:
            print(f'Error leyendo {file_path}: {e}')
            return None

    def run_merge(self):
        if not self.excel_files:
            messagebox.showinfo('Info', 'Debe seleccionar al menos un archivo F42 Excel.')
            return

        try:
            self.status_label.config(text='Procesando archivos...')
            self.generate_btn.config(state=tk.DISABLED)
            self.root.update()

            # Leer todos los archivos
            dfs = []
            for f in self.excel_files:
                df = self.read_f42_file(f)
                if df is not None:
                    dfs.append(df)
                    print(f'Leído {os.path.basename(f)}: {len(df)} filas')

            if not dfs:
                messagebox.showerror('Error', 'No se pudieron leer los archivos. Verifique el formato.')
                self.status_label.config(text='')
                self.generate_btn.config(state=tk.NORMAL)
                return

            # Concatenar todos los DataFrames
            # Obtener todas las columnas únicas
            all_columns = []
            for df in dfs:
                for c in df.columns:
                    if c not in all_columns:
                        all_columns.append(c)

            # Alinear todos los DataFrames con las mismas columnas
            aligned = []
            for df in dfs:
                aligned.append(df.reindex(columns=all_columns))

            # Concatenar
            result = pd.concat(aligned, ignore_index=True)

            # Guardar en archivo temporal
            temp_dir = tempfile.mkdtemp(prefix='f42_merge_')
            out_path = os.path.join(temp_dir, 'F42_MERGED.xlsx')
            result.to_excel(out_path, index=False, engine='openpyxl')
            
            self.output_path = out_path
            self.open_btn.config(state=tk.NORMAL)
            # Mostrar mensaje
            self.status_label.config(text='Archivo F42 merged generado')
            messagebox.showinfo('Éxito', 'Archivo F42 merged generado')
            
        except Exception as exc:
            messagebox.showerror('Error', f'Ocurrió un error al procesar: {exc}')
            self.status_label.config(text='Error durante el procesamiento')
        finally:
            self.generate_btn.config(state=tk.NORMAL)

    def open_output(self):
        if not self.output_path:
            return
        try:
            if os.name == 'nt':
                os.startfile(self.output_path)
            elif sys.platform == 'darwin':
                os.system(f'open "{self.output_path}"')
            else:
                os.system(f'xdg-open "{self.output_path}"')
        except Exception:
            messagebox.showinfo('Ubicación', f'Archivo generado en: {self.output_path}')


def main():
    root = tk.Tk()
    app = F42MergeApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
