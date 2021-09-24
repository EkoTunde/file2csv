import os
from pathlib import Path
from shipment import Shipment
import subprocess
import tkinter as tk
from tkinter import Button, Toplevel, filedialog, messagebox, Text, INSERT
from extractor import MLExtractor, TNExtractor


class Application(tk.Frame):

    FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

    def __init__(self, master):
        super().__init__(master)
        self.pack(padx=10, pady=10)
        self.selected_index = 0
        self.create_widgets()
        self.mercado_libre_converter = MLExtractor()
        self.tienda_nube_converter = TNExtractor()
        self.pdf_scanned = False

    def create_converter(self):
        """Creates an instance of converter class,
        which will be used for converting PDF file onto
        Shipment objects.
        """
        self.mercado_libre_converter = MLExtractor()
        self.tienda_nube_converter = TNExtractor()

    def create_widgets(self):
        """Creates and grid/pack all tkinter widget
        which will be rendered on screen.
        """

        # Just a label
        self.entry_file_label = tk.Label(
            self, text="Elegir archivo", padx=5, pady=10)
        self.entry_file_label.grid(row=0, column=0, sticky="w")

        # Entry for filepath
        self.entry_file = tk.Entry(self, width=60)
        self.entry_file.grid(row=0, column=1, sticky="ew")

        # Just space
        self.row_0_space = tk.Label(self, padx=1)
        self.row_0_space.grid(row=0, column=2)

        # Select file button
        self.entry_file_button = tk.Button(
            self, text="...", height=1, padx=10, command=self.select_file)
        self.entry_file_button.grid(row=0, column=3)

        # Scanning frame
        self.frame = tk.Frame(self, pady=5)
        self.frame.grid(row=1, column=0, columnspan=4, sticky="we")

        # Scan explanation label
        explanation = "Presioná \"Escanear\" para buscar los envíos en el pdf"
        self.status_label = tk.Label(self.frame, text=explanation)
        self.status_label.pack()

        # Scann for MercadoLibre data button
        self.entry_file_button = tk.Button(
            self.frame, text="Escanear ML", height=1, padx=10,
            command=self.convert)
        self.entry_file_button.pack()

        # Scann for TiendaNube data button
        self.entry_file_button = tk.Button(
            self.frame, text="Escanear TiendaNube", height=1, padx=10,
            command=self.convert_tienda_nube)
        self.entry_file_button.pack()

        # Scanning status label
        self.status_label = tk.Label(self.frame, text="En espera")
        self.status_label.pack()

        # Previewer frame
        self.previewer_frame = tk.Frame(self, pady=10)
        self.previewer_frame.grid(row=2, column=1, columnspan=4, sticky="we")

        # Label for shipment number
        self.shipment_no = tk.Label(
            self.previewer_frame, text="", padx=5, pady=2)
        self.shipment_no.grid(row=0, column=1, columnspan=2, sticky="ew")

        # Label for Cliente
        self.cliente_label = tk.Label(
            self.previewer_frame, text="Cliente", padx=5, pady=2)
        self.cliente_label.grid(row=1, column=0, sticky="w")

        # Entry for Cliente
        self.cliente_entry = tk.Entry(self.previewer_frame)
        self.cliente_entry.grid(row=1, column=1, columnspan=2, sticky="ew")

        # Label for ID-Cliente-Flex
        self.id_cliente_label = tk.Label(
            self.previewer_frame, text="ID-Cliente-Flex", padx=5, pady=2)
        self.id_cliente_label.grid(row=2, column=0, sticky="w")

        # Entry for ID-Cliente-Flex
        self.id_cliente_entry = tk.Entry(self.previewer_frame)
        self.id_cliente_entry.grid(row=2, column=1, sticky="ew", columnspan=2)

        # Label for Cantidad
        self.cantidad_label = tk.Label(
            self.previewer_frame, text="Cantidad", padx=5, pady=2)
        self.cantidad_label.grid(row=3, column=0, sticky="w")

        # Entry for Cantidad
        self.cantidad_entry = tk.Entry(self.previewer_frame)
        self.cantidad_entry.grid(row=3, column=1, sticky="ew", columnspan=2)

        # Label for Domicilio
        self.domicilio_label = tk.Label(
            self.previewer_frame, text="Domicilio", padx=5, pady=2)
        self.domicilio_label.grid(row=4, column=0, sticky="w")

        # Entry for Domicilio
        self.domicilio_entry = tk.Entry(self.previewer_frame)
        self.domicilio_entry.grid(row=4, column=1, sticky="ew", columnspan=2)

        # Label for Zona
        self.zona_label = tk.Label(
            self.previewer_frame, text="Zona", padx=5, pady=2)
        self.zona_label.grid(row=5, column=0, sticky="w")

        # Entry for Zona
        self.zona_entry = tk.Entry(self.previewer_frame)
        self.zona_entry.grid(row=5, column=1, sticky="ew", columnspan=2)

        # Label for Subzona
        self.subzona_label = tk.Label(
            self.previewer_frame, text="Subzona", padx=5, pady=2)
        self.subzona_label.grid(row=6, column=0, sticky="w")

        # Entry for Subzona
        self.subzona_entry = tk.Entry(self.previewer_frame)
        self.subzona_entry.grid(row=6, column=1, sticky="ew", columnspan=2)

        # Label for Codigo Postal
        self.codigo_postal_label = tk.Label(
            self.previewer_frame, text="Codigo Postal", padx=5, pady=2)
        self.codigo_postal_label.grid(row=7, column=0, sticky="w")

        # Entry for Codigo Postal
        self.codigo_postal_entry = tk.Entry(self.previewer_frame)
        self.codigo_postal_entry.grid(
            row=7, column=1, sticky="ew", columnspan=2)

        # Label for ID-Envío-Flex
        self.id_envio_label = tk.Label(
            self.previewer_frame, text="ID-Envío-Flex", padx=5, pady=2)
        self.id_envio_label.grid(row=8, column=0, sticky="w")

        # Entry for ID-Envío-Flex
        self.id_envio_entry = tk.Entry(self.previewer_frame)
        self.id_envio_entry.grid(row=8, column=1, sticky="ew", columnspan=2)

        # Label for ID-Venta
        self.id_venta_label = tk.Label(
            self.previewer_frame, text="ID-Venta", padx=5, pady=2)
        self.id_venta_label.grid(row=9, column=0, sticky="w")

        # Entry for ID-Venta
        self.id_venta_entry = tk.Entry(self.previewer_frame)
        self.id_venta_entry.grid(row=9, column=1, sticky="ew", columnspan=2)

        # Label for Destinatario
        self.destinatario_label = tk.Label(
            self.previewer_frame, text="Destinatario", padx=5, pady=2)
        self.destinatario_label.grid(row=10, column=0, sticky="w")

        # Entry for Destinatario
        self.destinatario_entry = tk.Entry(self.previewer_frame)
        self.destinatario_entry.grid(
            row=10, column=1, sticky="ew", columnspan=2)

        # Button Previous
        self.prev_btn = tk.Button(
            self.previewer_frame, text="<", width=3, command=self.prev_item)
        self.prev_btn.grid(row=11, column=0)

        # Button Save
        self.prev_btn = tk.Button(
            self.previewer_frame, text="Guardar", command=self.save_item)
        self.prev_btn.grid(row=11, column=1, sticky="we")

        # Button Next
        self.next_btn = tk.Button(
            self.previewer_frame, text=">", width=3, command=self.next_item)
        self.next_btn.grid(row=11, column=2, sticky="e")

        # Export frame
        self.export_frame = tk.LabelFrame(
            self, text="Exportación", pady=10, padx=10)
        self.export_frame.grid(row=3, column=0, columnspan=4, sticky="we")

        # Copy to clipboard
        self.copy_button = tk.Button(
            self.export_frame, text="Copiar", padx=15,
            command=lambda: self.copy_to_clipboard(self.shipments_as_text()))
        self.copy_button.pack(side=tk.LEFT)

        # Export to CSV
        explanation = "Exportar a CSV"
        self.export_csv = tk.Button(
            self.export_frame, text='Exportar a CSV', padx=15,
            command=lambda: self.export('csv'))
        self.export_csv.pack(side=tk.LEFT)

        # Export to TXT
        self.export_txt = tk.Button(
            self.export_frame, text='Exportar a TXT', padx=15,
            command=lambda: self.export('txt'))
        self.export_txt.pack(side=tk.LEFT)

        # Open export string as dialog
        self.export_dialog = tk.Button(
            self.export_frame, text='Ventana de diálogo', padx=15,
            command=self.open_as_dialog)
        self.export_dialog.pack(side=tk.LEFT)

        self.columnconfigure(1, weight=1)

        return

    def select_file(self):
        """Runs an Open File Dialog for PDF selecting
        """
        try:
            downloads_path = str(Path.home() / "Downloads")
            self.filename = filedialog.askopenfilename(
                # initialdir="C:/", title="Seleccioná un archivo PDF",
                initialdir=downloads_path, title="Seleccioná un archivo PDF",
                filetypes=(("PDF", "*.pdf"), ("all", "*,*"))
            )
            self.entry_file.delete(0, 'end')
            self.entry_file.insert(0, self.filename)
            self.show_in_hold()
            return
        except Exception as e:
            self.popup(e)

    def convert(self):
        """Performs the conversion from PDF to a list of
            objects representing the MercadoLibre's shipments.
        """
        if len(self.entry_file.get()) > 0:
            self.shipments = []
            self.clear_entries()
            try:
                self.show_in_progress()
                self.shipments = self.mercado_libre_converter.convert(
                    self.entry_file.get())
                self.max = len(self.shipments)-1
                self.publish_item()
                self.show_success()
                self.pdf_scanned = True
            except Exception as e:
                self.pdf_scanned = False
                self.show_error()
                self.popup(e)
        else:
            self.popup("Ningún archivo seleccionado para escanear.")

    def convert_tienda_nube(self):
        """Performs the conversion from PDF to a list of
            objects representing the TiendaNube's shipments.
        """
        if len(self.entry_file.get()) > 0:
            self.shipments = []
            self.clear_entries()
            try:
                self.show_in_progress()
                self.shipments = self.tienda_nube_converter.convert(
                    self.entry_file.get())
                self.max = len(self.shipments)-1
                self.publish_item()
                self.show_success()
                self.pdf_scanned = True
            except Exception as e:
                self.pdf_scanned = False
                self.show_error()
                self.popup(e)
        else:
            self.popup("Ningún archivo seleccionado para escanear.")

    def show_error(self):
        """Displays an error message
        """
        self.status_label.config(
            text='Error al obtener los datos', fg="#F00")

    def show_in_progress(self):
        """Displays PDF convertion in progress
        """
        self.status_label.config(text='Procesando...', fg="#000")

    def show_in_hold(self):
        """Displays PDF convertions is awaiting for user to
            run it.
        """
        self.status_label.config(text='En espera', fg="#000")

    def show_success(self):
        """Displays PDF has been successfully converted.
        """
        t = f'Éxito! {len(self.shipments)} envíos encontrados!'
        self.status_label.config(text=t, fg="#008000")

    def next_item(self):
        """Updates selected item index from shipents's list.

        Returns:
            A popup error if pdf hasn't been converted yet
        """
        if not self.pdf_scanned:
            return self.popup("Todavía no se escaneó ningún PDF")
        if self.selected_index == self.max:
            self.selected_index = 0
        else:
            self.selected_index += 1
        self.publish_item()
        return

    def prev_item(self):
        """Updates selected item index from shipment's list.

        Returns:
            A popup error if pdf hasn't been converted yet
        """
        if not self.pdf_scanned:
            return self.popup("Todavía no se escaneó ningún PDF")
        if self.selected_index == 0:
            self.selected_index = self.max
        else:
            self.selected_index -= 1
        self.publish_item()

    def publish_item(self):
        """Displays the shipment at the selected index from shipment's list.
        """
        self.shipment_no.config(text=f'Envío N.° {self.selected_index+1}')
        self.clear_entries()
        self.cliente_entry.insert(
            0, str(self.shipments[self.selected_index].client_name))
        self.id_cliente_entry.insert(
            0, str(self.shipments[self.selected_index].client_id))
        self.cantidad_entry.insert(
            0, str(self.shipments[self.selected_index].quantity))
        self.domicilio_entry.insert(
            0, str(self.shipments[self.selected_index].direccion))
        self.zona_entry.insert(
            0, str(self.shipments[self.selected_index].ciudad))
        self.subzona_entry.insert(
            0, str(self.shipments[self.selected_index].barrio))
        self.codigo_postal_entry.insert(
            0, str(self.shipments[self.selected_index].codigo_postal))
        self.id_envio_entry.insert(
            0, str(self.shipments[self.selected_index].id_envio))
        self.id_venta_entry.insert(
            0, str(self.shipments[self.selected_index].id_venta))
        self.destinatario_entry.insert(
            0, str(self.shipments[self.selected_index].destinatario))
        return

    def clear_entries(self):
        """Clears all Entry Widgets which indicate shipment's info.
        """
        for entry in self.get_entries():
            entry.delete(0, 'end')

    def save_item(self):
        """Saves the currently displayed shipment item info.

        Returns:
            A popup error if pdf hasn't been converted yet or a
            popup error displaying the exception.
        """
        if not self.pdf_scanned:
            return self.popup("Todavía no se escaneó ningún PDF")
        try:
            self.shipments[self.selected_index].client_name = \
                self.cliente_entry.get()
            self.shipments[self.selected_index].client_id = \
                self.id_cliente_entry.get()
            self.shipments[self.selected_index].quantity = \
                self.cantidad_entry.get()
            self.shipments[self.selected_index].direccion = \
                self.domicilio_entry.get()
            self.shipments[self.selected_index].ciudad = \
                self.zona_entry.get()
            self.shipments[self.selected_index].barrio = \
                self.subzona_entry.get()
            self.shipments[self.selected_index].codigo_postal = \
                self.codigo_postal_entry.get()
            self.shipments[self.selected_index].id_envio = \
                self.id_envio_entry.get()
            self.shipments[self.selected_index].id_venta = \
                self.id_venta_entry.get()
            self.shipments[self.selected_index].destinatario = \
                self.destinatario_entry.get()
            return
        except Exception as e:
            self.popup(e)

    def popup(self, message):
        """Displays error popup with the message provided
            parsed to string.

        Args:
            message (Any)
        """
        error_title = "Ha ocurrido un error"
        error_message = "Error: " + str(message)
        messagebox.showerror(error_title, error_message)
        return

    def export(self, extension: str):
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
        if not self.pdf_scanned:
            return self.popup("Todavía no se escaneó ningún PDF")
        try:
            file = filedialog.asksaveasfile(
                mode='w', defaultextension=f'*.{extension}',
                filetypes=((extension.upper(), f'*.{extension}'),
                           ("Todos los archivos", "*.*")))

            # asksaveasfile return `None` if dialog closed with "cancel".
            if file is None:
                return

            titles = "client_name,client_id,quantity,direccion," + \
                "zona,subzona,codigo_postal,id_envio,id_venta,destinatario\n"
            file.write(titles)
            file.write(self.shipments_as_text())
            file.close()
            name = file.name
            last_slash_index = name.rfind('/')
            path = name[:last_slash_index]
            return self.explore(path)
        except Exception as e:
            self.popup(e)

    def shipments_as_text(self):
        """Returns a joined string of all shipments previously parsed
        to CSV-style str.

        Returns:
            [type]: [description]
        """
        result = "".join([self.shipment_to_csv_string(shipment)
                          for shipment in self.shipments])
        return result[:-1]

    def open_as_dialog(self):
        """Displays all shipments parsed to full str in CSV-style
        in a new tkinter window.

        Returns:
            A popup error if pdf hasn't been converted yet
        """
        if not self.pdf_scanned:
            return self.popup("Todavía no se escaneó ningún PDF")
        top = Toplevel()
        my_label = Text(top)
        my_label.insert(INSERT, self.shipments_as_text())
        my_label.pack()
        btn = Button(top, text="Copiar",
                     command=lambda: self.copy_to_clipboard(
                         self.shipments_as_text()))
        btn.pack()

    def copy_to_clipboard(self, txt: str):
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

    def shipment_to_csv_string(self, shipment: Shipment) -> str:
        """Parses a shipment to a full str in CSV-style.

        Args:
            shipment (Shipment): to be parsed.

        Returns:
            str: shipment in CSV-style.
        """
        return \
            str(shipment.client_name).replace(", ", "") + "," + \
            str(shipment.client_id).replace(", ", "") + "," + \
            str(shipment.quantity).replace(", ", "") + "," + \
            str(shipment.direccion).replace(", ", "") + "," + \
            str(shipment.ciudad).replace(", ", "") + "," + \
            str(shipment.barrio).replace(", ", "") + "," + \
            str(shipment.codigo_postal).replace(", ", "") + "," + \
            str(shipment.id_envio).replace(", ", "") + "," + \
            str(shipment.id_venta).replace(", ", "") + "," + \
            str(shipment.destinatario).replace(", ", "") + \
            "\n"

    def explore(self, path: str):
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

    def get_entries(self):
        """Returns list containg all entry widgets
        which display a shipment fields.

        Returns:
            list: of entry widgets.
        """
        return [
            self.cliente_entry,
            self.id_cliente_entry,
            self.cantidad_entry,
            self.domicilio_entry,
            self.zona_entry,
            self.subzona_entry,
            self.codigo_postal_entry,
            self.id_envio_entry,
            self.id_venta_entry,
            self.destinatario_entry,
        ]


def main():
    root = tk.Tk()
    root.geometry("650x535")
    root.title("PDF2CSV")
    root.resizable(False, False)
    path = os.path.abspath(os.getcwd())
    icon = path + "\\img\\pdf2csv.gif"
    img = tk.PhotoImage(file=icon)
    root.tk.call('wm', 'iconphoto', root._w, img)
    app = Application(root)
    app.mainloop()


if __name__ == "__main__":
    main()
