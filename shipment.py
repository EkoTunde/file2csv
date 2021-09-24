from dataclasses import dataclass


@dataclass
class Shipment:
    quantity: int = None
    id_venta: int = None
    id_envio: int = None
    ciudad: str = None
    barrio: str = None
    direccion: str = None
    codigo_postal: int = None
    destinatario: str = None
    client_name: str = None
    client_id: int = None


@dataclass
class Shipment2:
    traking_id: str = ""
    domicilio: str = ""
    entrecalles: str = ""
    codigo_postal: str = ""
    localidad: str = ""
    partido: str = ""
    destinatario: str = ""
    dni_destinatario: str = ""
    detalle_envio: str = ""
    phone: str = ""
