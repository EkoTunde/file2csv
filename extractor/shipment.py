import unidecode
from dataclasses import dataclass


@dataclass
class Shipment:
    traking_id: str = ""
    domicilio: str = ""
    referencia: str = ""
    codigo_postal: str = ""
    localidad: str = ""
    partido: str = ""
    destinatario: str = ""
    dni_destinatario: str = ""
    phone: str = ""
    detalle_envio: str = ""

    def clean(self):
        self.replace_commas()
        self.replace_accents()
        return

    def replace_commas(self):
        self.traking_id = str(self.traking_id).replace(",", " ")
        self.domicilio = str(self.domicilio).replace(",", " ")
        self.referencia = str(self.referencia).replace(",", " ")
        self.codigo_postal = str(self.codigo_postal).replace(",", " ")
        self.localidad = str(self.localidad).replace(",", " ")
        self.partido = str(self.partido).replace(",", " ")
        self.destinatario = str(self.destinatario).replace(",", " ")
        self.dni_destinatario = str(self.dni_destinatario).replace(",", " ")
        self.phone = str(self.phone).replace(",", " ")
        self.detalle_envio = str(self.detalle_envio).replace(",", " ")
        return

    def replace_accents(self):
        self.traking_id = unidecode.unidecode(self.traking_id)
        self.domicilio = unidecode.unidecode(self.domicilio)
        self.referencia = unidecode.unidecode(self.referencia)
        self.codigo_postal = unidecode.unidecode(self.codigo_postal)
        self.localidad = unidecode.unidecode(self.localidad)
        self.partido = unidecode.unidecode(self.partido)
        self.destinatario = unidecode.unidecode(self.destinatario)
        self.dni_destinatario = unidecode.unidecode(self.dni_destinatario)
        self.phone = unidecode.unidecode(self.phone)
        self.detalle_envio = unidecode.unidecode(self.detalle_envio)
        return

    def is_not_empty(self):
        return len(self.traking_id) > 0 or len(self.domicilio) > 0 \
            or len(self.referencia) > 0 \
            or len(self.codigo_postal) > 0 \
            or len(self.localidad) > 0 \
            or len(self.partido) > 0 \
            or len(self.destinatario) > 0 \
            or len(self.dni_destinatario) > 0 \
            or len(self.phone) > 0 \
            or len(self.detalle_envio) > 0
