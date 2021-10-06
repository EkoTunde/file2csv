import os
import webbrowser
from pathlib import Path
import subprocess
import tkinter as tk
from tkinter import Button, Toplevel, filedialog, messagebox, Text, INSERT
from extractor.extractor import Extractor


class Application(tk.Frame):

    FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

    def __init__(self, master):
        super().__init__(master)
        self.pack(padx=10, pady=10)
        self.selected_index = 0
        self.create_widgets()
        self.extractor = Extractor()
        self.result = None
        self.shipments = []
        self.file_scanned = False

    def create_widgets(self):
        """Creates and grid/pack all tkinter widget
        which will be rendered on screen.
        """
        # Scanning frame
        self.frame_presentation = tk.Frame(self, pady=5)
        self.frame_presentation.grid(
            row=0, column=0, columnspan=3, sticky="we")

        # Presentation label, explaining program
        presentation_label_txt = "Escáner de archivos PDF " + \
            "de MercadoLibre o TiendaNube, o XLSX (Excel) " + \
            "con formato específico."
        self.presentation_label = tk.Label(
            self.frame_presentation,
            text=presentation_label_txt, padx=5, pady=10)
        self.presentation_label.pack()

        # Open Excel instructions.
        self.excel_instructions = tk.Button(
            self.frame_presentation, text='Ayuda sobre Excel', padx=15,
            command=self.__open_excel_instructions)
        self.excel_instructions.pack(side=tk.LEFT)

        # Scanning frame
        self.frame_0 = tk.Frame(self, pady=5)
        self.frame_0.grid(row=1, column=0, columnspan=3, sticky="we")

        # Just a label
        self.entry_file_label = tk.Label(
            self.frame_0, text="Elegir archivo", padx=5, pady=10)
        self.entry_file_label.grid(row=0, column=0, sticky="w")

        # Entry for filepath
        self.entry_file = tk.Entry(self.frame_0, width=60)
        self.entry_file.grid(row=0, column=1, sticky="ew")

        # Just space
        self.row_0_space = tk.Label(self.frame_0, padx=1)
        self.row_0_space.grid(row=0, column=2)

        # Select file button
        self.entry_file_button = tk.Button(
            self.frame_0, text="...", height=1,
            padx=10, command=self.__select_file)
        self.entry_file_button.grid(row=0, column=3)

        # Scanning frame
        self.frame = tk.Frame(self, pady=5)
        self.frame.grid(row=2, column=0, columnspan=4, sticky="we")

        # Scan explanation label
        explanation = "Presioná \"Escanear\" para buscar los" + \
            " envíos en el archivo"
        self.status_label = tk.Label(self.frame, text=explanation)
        self.status_label.pack()

        # Scann for Shipments' data button
        self.entry_file_button = tk.Button(
            self.frame, text="Escanear", height=1, padx=10,
            command=self.__convert)
        self.entry_file_button.pack()

        # Scanning status label
        self.status_label = tk.Label(self.frame, text="En espera")
        self.status_label.pack()

        # Export frame
        self.export_frame = tk.LabelFrame(
            self, text="Exportación", pady=10, padx=10)
        self.export_frame.grid(row=3, column=0, columnspan=3)

        # Copy to clipboard
        self.copy_button = tk.Button(
            self.export_frame, text="Copiar", padx=15,
            command=lambda: self.__copy_to_clipboard(self.result))
        self.copy_button.pack(side=tk.LEFT)

        # Export to CSV
        self.export_csv = tk.Button(
            self.export_frame, text='Exportar a CSV', padx=15,
            command=lambda: self.__export('csv'))
        self.export_csv.pack(side=tk.LEFT)

        # # Export to TXT
        self.export_txt = tk.Button(
            self.export_frame, text='Exportar a TXT', padx=15,
            command=lambda: self.__export('txt'))
        self.export_txt.pack(side=tk.LEFT)

        # Open export string as dialog
        self.export_dialog = tk.Button(
            self.export_frame, text='Ventana de diálogo', padx=15,
            command=self.__open_as_dialog)
        self.export_dialog.pack(side=tk.LEFT)

        # self.columnconfigure(1, weight=1)
        # self.columnconfigure(0, weight=1)

        self.__disable_buttons()
        return

    def __disable_buttons(self):
        self.all_buttons = [self.copy_button,
                            self.export_csv,
                            self.export_txt,
                            self.export_dialog]
        for btn in self.all_buttons:
            btn['state'] = 'disabled'
        return

    def __enable_buttons(self):
        self.all_buttons = [self.copy_button,
                            self.export_csv,
                            self.export_txt,
                            self.export_dialog]
        for btn in self.all_buttons:
            btn['state'] = 'normal'
        return

    def __select_file(self):
        """Runs an Open File Dialog for PDF selecting
        """
        try:
            downloads_path = str(Path.home() / "Downloads")
            self.filename = filedialog.askopenfilename(
                # initialdir="C:/", title="Seleccioná un archivo PDF",
                initialdir=downloads_path, title="Seleccioná un archivo PDF",
                filetypes=(("PDF, CSV, Excel", "*.*"),)
            )
            if self.filename:
                self.entry_file.delete(0, 'end')
                self.entry_file.insert(0, self.filename)
                return
        except Exception as e:
            self.__disable_buttons()
            self.__popup(e)

    def __convert(self):
        """Performs the conversion from file to a csv string
        representing the MercadoLibre's shipments.
        """
        if len(self.entry_file.get()) > 0:
            try:
                self.__disable_buttons()
                self.__show_in_progress()
                extraction = self.extractor.get_shipments(
                    self.entry_file.get())
                self.result, self.total_count = extraction
                self.file_scanned = True
                self.__show_success()
                self.__enable_buttons()

            except Exception as e:
                self.__show_error()
                return self.__popup(e)
        else:
            return self.__popup("Ningún archivo seleccionado para escanear.")

    def __show_error(self):
        """Displays an error message
        """
        self.status_label.config(
            text='Error al obtener los datos', fg="#F00")
        return

    def __show_in_progress(self):
        """Displays File convertion in progress
        """
        self.status_label.config(text='Procesando...', fg="#000")
        return

    def __show_success(self):
        """Displays File has been successfully converted.
        """
        t = f'Éxito! {self.total_count} envíos encontrados!'
        self.status_label.config(text=t, fg="#008000")
        return

    def __popup(self, message):
        """Displays error popup with the message provided
            parsed to string.

        Args:
            message (Any)
        """
        error_title = "Ha ocurrido un error"
        error_message = "Error: " + str(message)
        messagebox.showerror(error_title, error_message)
        return

    def __export(self, extension: str):
        """Performs the export to a file in the provided extension with shipments
            formatted in CSV-style.
            Asks for filename and path to save to.

        Args:
            extension (str): file extension (csv, txt).

        Returns:
            A popup error if pdf hasn't been converted yet, or a
            popup error displaying the exception, or nothing
            if there wasn't any filepath to save to.
        """
        if not self.file_scanned:
            return self.__popup("Todavía no se escaneó ningún PDF")
        try:
            file = filedialog.asksaveasfile(
                mode='w', defaultextension=f'*.{extension}',
                filetypes=((extension.upper(), f'*.{extension}'),
                           ("Todos los archivos", "*.*")))

            # asksaveasfile return `None` if dialog closed with "cancel".
            if file is None:
                return

            file.write(self.result)
            file.close()
            name = file.name
            last_slash_index = name.rfind('/')
            path = name[:last_slash_index]
            return self.__explore(path)
        except Exception as e:
            self.__popup(e)

    def __open_excel_instructions(self):
        path = os.path.abspath(os.getcwd())
        filename = path + "\\excel_instructions.html"
        return webbrowser.open('file://' + os.path.realpath(filename))

    def __open_txt_as_dialog(self, txt, big=None):
        """Displays text in a new tkinter window.
        """
        top = Toplevel()
        my_label = Text(top)
        my_label.insert(INSERT, txt)
        my_label.pack()
        if big:
            top.resizable(False, False)
        return top

    def __open_as_dialog(self):
        """Displays all shipments parsed to full str in CSV-style
        in a new tkinter window.

        Returns:
            A popup error if pdf hasn't been converted yet
        """
        if not self.file_scanned:
            return self.__popup("Todavía no se escaneó ningún PDF")
        top = self.__open_txt_as_dialog(self.result)
        btn = Button(top, text="Copiar",
                     command=lambda: self.__copy_to_clipboard(
                         self.result))
        btn.pack()
        return

    def __copy_to_clipboard(self, txt: str):
        """Copies to clipboard displayed shipments parsed to full str
        in CSV-style displayed in new tkinter window.

        Args:
            txt (Str): Text to copy.
        """
        self.clipboard_clear()
        self.clipboard_append(txt)
        # now it stays on the clipboard after the window is closed
        self.update()
        return

    def __explore(self, path: str):
        """Opens File Explorer at provided path.

        Args:
            path (str): to open File Explorer to.
        """
        # explorer would choke on forward slashes
        path = os.path.normpath(path)
        if os.path.isdir(path):
            subprocess.run([self.FILEBROWSER_PATH, path])
        elif os.path.isfile(path):
            subprocess.run(
                [self.FILEBROWSER_PATH, '/select,', os.path.normpath(path)])


def main():
    try:
        root = tk.Tk()
        root.geometry("580x330")
        root.title("File2CSV")
        root.resizable(False, False)
        path = os.path.abspath(os.getcwd())
        icon = path + "\\file2csv.gif"
        img = tk.PhotoImage(file=icon)
        root.tk.call('wm', 'iconphoto', root._w, img)
        app = Application(root)
        app.mainloop()
    except Exception as e:
        logpath = os.path.abspath(os.getcwd()) + "\\log.txt"
        with open(logpath, 'a') as txt:
            txt.write(str(e) + "\n")
    return


if __name__ == "__main__":
    main()
