import fitz  # this is pymupdf
import openpyxl
from shipment import Shipment
from dataclasses import asdict


class PDFExtractor(object):

    def __init__(self, filepath=None, **kwargs):
        self.pdf = filepath
        self.SEPARATOR = "~"*6
        self.text = ""
        super().__init__(**kwargs)

    def convert(self, pdf=None) -> list:
        """Initializes convertion of pdf to dict called 'shipments'

        Args:
            pdf (str, optional): PDF file name. Defaults to None.

        Returns:
            list: containing 'shipments' dicts
        """
        if pdf:
            self.pdf = pdf
        text = self.get_text()
        self.shipments = self.get_shipments(text)
        return self.shipments

    def parse_clients(self, extra_info):
        clients = []
        if extra_info:
            for line in extra_info.split("\n"):
                if "#" in line:
                    parts = line.split(" #")
                    clients.append({
                        'client_name': parts[0],
                        'client_id': parts[1],
                    })
            return clients
        return None

    def create_csv(self, pathtosave: str):
        with open(pathtosave, "w") as csv:
            titles = "quantity, id_venta, id_envio, ciudad, barrio, " + \
                "direccion, codigo_postal, destinatario," + \
                " client_name, client_id\n"
            csv.write(titles)
            for shipment in self.shipments:
                result = ""
                for key, value in asdict(shipment).items():
                    result += str(value).replace(",", "") + ","
                result = result[:-2] + "\n"
                csv.write(result)


class MLExtractor(PDFExtractor):

    def __init__(self, filepath=None, **kwargs):
        super().__init__(filepath, **kwargs)

    def get_text(self):
        """Opens the PDF file and extracts it as a single str.

        Returns:
            str: containing the whole pdf file as text.
        """
        text = ""
        with fitz.open(self.pdf) as doc:
            for page in doc:
                extract = page.getText()
                print(extract)

                # Keeps track of shipments
                # Which are max 3 per page
                counter = 0

                # This page contains the final list
                if "Despacha" in extract:
                    break

                for line in extract.split("\n"):

                    # If line is clean, append it
                    if self.is_clean_line(line):
                        text += line

                        # End of column was reached
                        if "CP:" in line and counter < 3:
                            text += "\n"
                            text += self.SEPARATOR
                            counter += 1

                        text += '\n'
        return text

    def is_clean_line(self, line):
        not_flex = "Flex" not in line
        not_recorta = "Recorta" not in line
        not_nickname = "Nickname" not in line
        return not_flex and not_recorta and not_nickname

    def get_shipments(self, text):
        pages = text.split("\n\n")
        shipments = []
        temp = 1
        for page in pages:
            parts = page.split(self.SEPARATOR+"\n")
            if temp == 1:
                temp = 0
            q = len(parts)
            if q != 1:
                ship_1 = self.parse_package(parts[0]) if q > 1 else None
                ship_2 = self.parse_package(parts[1]) if q > 2 else None
                ship_3 = self.parse_package(parts[2]) if q > 3 else None
                extra_info = parts[-1]
                clients = self.parse_clients(extra_info)
                for i, ship in enumerate([ship_1, ship_2, ship_3]):
                    if ship:
                        ship.client_name = clients[i]['client_name']
                        ship.client_id = clients[i]['client_id']
                        shipments.append(ship)
        return shipments

    def parse_package(self, input: str):
        lines = input.split("\n")
        shipment = Shipment()
        if lines[0] == "":
            return None
        for line in reversed(lines):
            if "CP" in line:
                shipment.codigo_postal = self.parse_without(
                    line, "CP: ", as_int=True)
                break
        if lines[0].isnumeric():    # Cuando el pedido es solo un producto
            shipment.quantity = int(lines[0])
            shipment.id_venta = self.parse_without(
                lines[3], "Venta: ", as_int=True)
            shipment.id_envio = self.parse_without(
                lines[4], "Tracking: ", as_int=True)
            shipment.ciudad = lines[5]
            shipment.barrio = lines[6]
            shipment.direccion = self.parse_without(lines[8], "Direccion: ")
            shipment.destinatario = self.parse_without(
                lines[7], "Destinatario: ")
        else:
            shipment.quantity = int(lines[2])
            shipment.id_venta = self.parse_without(
                lines[0], "Venta: ", as_int=True)
            shipment.id_envio = self.parse_without(
                lines[1], "Tracking: ", as_int=True)
            shipment.ciudad = lines[4]
            shipment.barrio = lines[5]
            shipment.direccion = self.parse_without(lines[7], "Direccion: ")
            shipment.destinatario = self.parse_without(
                lines[6], "Destinatario: ")
        return shipment

    def parse_without(self, s: str, replaceable: str, as_int=False):
        cleaned = s.replace(replaceable, "")
        if as_int:
            cleaned = int(cleaned.replace(" ", ""))
        return cleaned


class TNExtractor(PDFExtractor):

    def __init__(self, filepath=None):
        super().__init__(filepath)

    def get_text(self):
        """Opens the PDF file and extracts it as a single str.

        Returns:
            str: containing the whole pdf file as text.
        """
        text = ""
        with fitz.open(self.pdf) as doc:
            for page in doc:
                extract = page.getText()
                text += str(extract).replace('\nOrden: #',
                                             f'\n{self.SEPARATOR}\nOrden: #')
                text += "\n\n"
        return text

    def get_shipments(self, text):
        pages = text.split("\n\n")
        shipments = []
        for page in pages:
            parts = page.split(self.SEPARATOR)
            q = len(parts)
            if q != 1:
                for i in range(q):
                    shipment = self.parse_package(parts[i])
                    shipments.append(shipment)
        return shipments

    def parse_package(self, input: str):
        lines = input.replace(" ,", ",").split('\n')
        if lines[0] == '':
            lines.pop(0)
        shipment = Shipment()
        shipment.id_venta = int(lines[0].split('#')[1])
        shipment.destinatario = lines[3]
        shipment.direccion = lines[4]
        if 'Teléfono: ' in lines[7]:
            lines.insert(5, lines[5].split(', ')[0])
        shipment.barrio = lines[5]
        location = lines[6].split(', ')
        shipment.ciudad = location[0] if len(location) > 0 else None
        shipment.codigo_postal = location[2].replace(
            ',', '') if len(location) > 2 else None
        shipment.client_name = lines[10]
        openning_index = input.index("Subtotal (")
        closing_index = input.index(" productos)")
        shipment.quantity = int(input[openning_index+10:closing_index])
        return shipment

    def parse_without(self, s: str, replaceable: str, as_int=False):
        cleaned = s.replace(replaceable, "")
        if as_int:
            cleaned = int(cleaned.replace(" ", ""))
        return cleaned


class ExcelExtractor(object):

    def convert(self, excel=None) -> list:
        """Initializes convertion of excel to dict called 'shipments'

        Args:
            excel (str, optional): Excel file name. Defaults to None.

        Returns:
            list: containing 'shipments' dicts
        """
        if excel:
            self.excel = excel
        text = self.get_text()
        self.shipments = self.get_shipments(text)
        return self.shipments

    # def get_text(self):
    #     wb = openpyxl.load_workbook('example.xlsx')
    #     pass

    # def __init__(self, filepath=None):
    #     super().__init__(filepath)

    # def get_text(self):
    #     """Opens the PDF file and extracts it as a single str.

    #     Returns:
    #         str: containing the whole pdf file as text.
    #     """
    #     text = ""
    #     with fitz.open(self.pdf) as doc:
    #         for page in doc:
    #             extract = page.getText()
    #             text += str(extract).replace('\nOrden: #',
    #                                          f'\n{self.SEPARATOR}\nOrden: #')
    #             text += "\n\n"
    #     return text


algo = """TRACKING ID MERCADO LIBRE	DOMICILIO	ENTRECALLES	CODIGO POSTAL	LOCALIDAD	PARTIDO	DNI_DESTINATARIO	DESTINATARIO	DETALLE DEL ENVIO"""


# 1) Ver qué formato es el archivo
# 2) Si es formato excel: convertir a csv
# 3) Si es formato csv: devolver csv
# 2) Si es pdf: verificar si es Meli o TiendaNube


class CantConvertNothing(Exception):
    pass


class FileWithNoExtension(Exception):
    pass


class InvalidExtension(Exception):
    pass


class ShipmentExtractor(object):

    def __init__(self, filepath: str = None):
        if filepath and type(filepath) != str:
            raise TypeError("Filepath must be a str.")
        self.file = filepath

    def convert(self, filepath: str = None) -> str:
        """Initializes convertion of file to csv

        Args:
            filepath (str, optional): file name. Defaults to None.

        Returns:
            csv: containing shipments's data
        """

        # If filepath was provided here, update self.file
        if filepath:
            # If filepath's not a str
            if type(filepath) != str:
                raise TypeError("Filepath must be a str.")
            self.file = filepath

        # If no filepath was provided
        if not self.file:
            raise CantConvertNothing("No filepath was provided.")

        try:
            # Find last '.' (stop)
            last_stop = self.file.rindex(".")
            # Get the extension
            extension = self.file[last_stop+1:]
        except ValueError:
            raise FileWithNoExtension("File didn't contain an extension.")

        if extension == 'pdf':
            # Do pdf work
            self.convert_pdf()
            pass
        elif extension == 'xlsx':
            # Do xlsx work
            print("do xlsx")
            pass
        elif extension == 'csv':
            # Do csv work
            print("do csv")
            pass
        else:
            raise InvalidExtension("Unknown file extension.")

        return 1

    def convert_pdf(self):
        with fitz.open(self.file) as pdf:
            first_page = pdf[0]
            if self.is_mercado_libre(first_page.getText()):
                self.convert_mercado_libre(pdf)
            else:
                self.convert_tienda_nube(pdf)
        return ""

    def is_mercado_libre(self, text):
        return 'Mercado Envíos' in text \
            or 'Flex' in text or 'Recorta' in text

    def convert_mercado_libre(self, pdf):
        extractor = MLExtractor(pdf)
        shipments = extractor.convert()
        print(shipments)
        return

    def convert_tienda_nube(self, pdf):
        pass

    def convert_xlsx(self):
        return ""

    def convert_csv(self):
        return ""
