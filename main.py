import os
from pathlib import Path
from shipment import Shipment
import subprocess
import tkinter as tk
from tkinter import filedialog
from extractor import PDF2CSV


class Application(tk.Frame):

    FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

    def __init__(self, master):
        super().__init__(master)
        self.pack(padx=10, pady=10)
        self.selected_index = 0
        self.create_widgets()
        self.create_converter()

    def create_converter(self):
        self.converter = PDF2CSV()

    def create_widgets(self):

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

        # Scann for data button
        self.entry_file_button = tk.Button(
            self.frame, text="Escanear", height=1, padx=10,
            command=self.convert)
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

        # Export to CSV
        self.export_btn = tk.Button(
            self, text="Exportar a CSV", command=self.export_to_csv)
        self.export_btn.grid(row=3, column=0, columnspan=4, sticky="we")

        self.columnconfigure(1, weight=1)

        return

    def select_file(self):
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

    def convert(self):
        self.shipments = []
        self.clear_entries()
        try:
            self.show_in_progress()
            self.shipments = self.converter.convert(self.entry_file.get())
            self.max = len(self.shipments)-1
            self.publish_item()
            self.show_success()
        except Exception as e:
            print(e)
            self.show_error()

    def show_error(self):
        self.status_label.config(
            text='Error al obtener los datos', fg="#F00")

    def show_in_progress(self):
        self.status_label.config(text='Procesando...', fg="#000")

    def show_in_hold(self):
        self.status_label.config(text='En espera', fg="#000")

    def show_success(self):
        t = f'Éxito! {len(self.shipments)} envíos encontrados!'
        self.status_label.config(text=t, fg="#008000")

    def next_item(self):
        if self.selected_index == self.max:
            self.selected_index = 0
        else:
            self.selected_index += 1
        self.publish_item()

    def prev_item(self):
        if self.selected_index == 0:
            self.selected_index = self.max
        else:
            self.selected_index -= 1
        self.publish_item()

    def publish_item(self):
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
        for entry in self.get_entries():
            entry.delete(0, 'end')

    def save_item(self):
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

    def export_to_csv(self):
        file = filedialog.asksaveasfile(mode='w', defaultextension=".csv")

        # asksaveasfile return `None` if dialog closed with "cancel".
        if file is None:
            return

        titles = "client_name,client_id,quantity,direccion," + \
            "zona,subzona,codigo_postal,id_envio,id_venta,destinatario\n"
        file.write(titles)
        for shipment in self.shipments:
            result = self.shipment_to_csv_string(shipment)
            file.write(result)
        file.close()
        name = file.name
        last_slash_index = name.rfind('/')
        path = name[:last_slash_index]
        return self.explore(path)

    def shipment_to_csv_string(self, shipment: Shipment) -> str:
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

    def explore(self, path):
        # explorer would choke on forward slashes
        path = os.path.normpath(path)
        if os.path.isdir(path):
            subprocess.run([self.FILEBROWSER_PATH, path])
        elif os.path.isfile(path):
            subprocess.run(
                [self.FILEBROWSER_PATH, '/select,', os.path.normpath(path)])

    def get_entries(self):
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
