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
