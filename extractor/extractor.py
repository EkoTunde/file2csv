from .exceptions import (
    CantConvertNothing,
    FileWithNoExtension,
    InvalidExtension
)
from .pdf_module import PDFModule
from .excel_module import ExcelModule
from .csv_module import CSVModule
from .shipment import Shipment


class Extractor():

    def __init__(self, filepath: str = None, **kwargs):
        if filepath and type(filepath) != str:
            raise TypeError("Filepath must be a str.")
        self.filepath = filepath

    def get_shipments(self, filepath: str = None) -> tuple:
        """Initializes convertion of file to csv

        Args:
            filepath (str, optional): file name. Defaults to None.

        Raises:
            TypeError: if given filepath isn't a string.
            CantConvertNothing: if there wasn't any filepath
            provided (if it's none).
            FileWithNoExtension: if the file doesn't have an extension.

        Returns:
            tuple[str, int]: tuple of csv string containing shipments's data
             and shipments count
        """
        # If filepath was provided here, update self.file
        if filepath:
            # If filepath's not a str
            if type(filepath) != str:
                raise TypeError("Filepath must be a str.")
            self.filepath = filepath

        # If no filepath was provided
        if not self.filepath:
            raise CantConvertNothing("No filepath was provided.")

        try:
            # Find last '.' (stop)
            last_stop = self.filepath.rindex(".")
            # Get the extension
            extension = self.filepath[last_stop+1:]
            # Extract the shipments
            shipments = self.__do_extraction(extension)
            # Return the shipments parsed as a csv
            return (self.__shipments_to_csv(shipments), len(shipments))
        except ValueError as e:
            print(e)
            # The file doesn't have an extension.
            raise FileWithNoExtension("File didn't contain an extension.")

    def __do_extraction(self, extension: str) -> list:
        """Extract the shipments from the file provided,
        for the specific given extension (pdf, xlsx, csv).

        Args:
            extension (str): a file extension expressed as a string.
            Compatible are: pdf, xlsx and csv. Pdf must be a file containing
            shipment labels from MercadoLibre and TiendaNube.

        Raises:
            InvalidExtension: for incompatible extensions
            (anything but pdf, xlsx & csv)

        Returns:
            list[Shipment2]: all the shipments found.
        """
        extraction_dict = {
            'pdf': PDFModule(self.filepath),
            'xlsx': ExcelModule(self.filepath),
            'csv': CSVModule(self.filepath),
        }
        try:
            return extraction_dict[extension].extract_shipments()
        except KeyError:
            raise InvalidExtension("Unknown file extension.")

    def __shipments_to_csv(self, shipments: list) -> str:
        """Parses all shipments to a csv string and adds the headers.

        Args:
            shipments (list[Shipment]): the shipments to parse.

        Returns:
            str: the csv string
        """
        titles = "traking_id,domicilio,referencia," + \
            "codigo_postal,localidad,partido,destinatario," + \
            "dni_destinatario,telefono,detalle_envio"
        shipments_mapped = list(map(
            self.__shipment_as_csv_row_string, shipments))
        shipments_str = "\n".join(shipments_mapped)
        return f'{titles}\n{shipments_str}'

    def __shipment_as_csv_row_string(self, shipment: Shipment) -> str:
        """Parses a Shipment object into a str as in a csv file's row.

        Args:
            shipment (Shipment): the object to map to string.

        Returns:
            str: the corresponding csv file's row.
        """
        values = [
            shipment.tracking_id, shipment.domicilio, shipment.referencia,
            shipment.codigo_postal, shipment.localidad, shipment.partido,
            shipment.destinatario, shipment.dni_destinatario,
            shipment.phone, shipment.detalle_envio,
        ]
        return ",".join(map(str, values))
