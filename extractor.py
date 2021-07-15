import fitz  # this is pymupdf
from shipment import Shipment
from dataclasses import asdict


class MLExtractor(object):

    SEPARATOR = "~"*6
    text = ""

    def __init__(self, filepath=None):
        self.pdf = filepath

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

    def get_text(self):
        """Opens the PDF file and extracts it as a single str.

        Returns:
            str: containing the whole pdf file as text.
        """
        text = ""
        with fitz.open(self.pdf) as doc:
            for page in doc:
                extract = page.getText()

                # Keeps track of shipments
                # Which are max 3 per page
                counter = 0

                # print(extract)

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
        # print(text)
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
            # print('part 0 =>' + parts[0]) if q > 1 else None
            # print('part 1 =>' + parts[1]) if q > 2 else None
            # print('part 2 =>' + parts[2]) if q > 3 else None
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
        print(input)
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


class TNExtractor(object):

    SEPARATOR = "~"*6
    text = ""

    def __init__(self, filepath=None):
        self.pdf = filepath

    def convert(self, pdf=None) -> list:
        """Initializes convertion of pdf to dict called 'shipments'

        Args:
            pdf (str, optional): PDF file name. Defaults to None.

        Returns:
            list: containing 'shipments' dicts
        """
        if pdf:
            self.pdf = pdf
        self.text = self.get_text()
        with open('prueba_tienda_mia.txt', 'a') as doc:
            doc.write(self.text)
        self.shipments = self.get_shipments(self.text)
        return self.shipments

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
                # ship_1 = self.parse_package(parts[0]) if q > 1 else None
                # ship_2 = self.parse_package(parts[1]) if q > 2 else None
                # ship_3 = self.parse_package(parts[2]) if q > 3 else None
                # ship_4 = self.parse_package(parts[3]) if q > 4 else None
                # extra_info = parts[-1]
                # clients = self.parse_clients(extra_info)
                # for i, ship in enumerate([ship_1, ship_2, ship_3, ship_4]):
                #     if ship:
                #         ship.client_name = clients[i]['client_name']
                #         ship.client_id = clients[i]['client_id']
                # shipments.append(ship_1)
                # shipments.append(ship_2)
                # shipments.append(ship_3)
                # shipments.append(ship_4)
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
