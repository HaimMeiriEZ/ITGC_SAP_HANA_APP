import os
from dataclasses import dataclass


@dataclass
class AppConfig:
    sap_system: str
    hana_host: str
    hana_port: int
    hana_user: str
    hana_password: str

    @classmethod
    def from_env(cls) -> "AppConfig":
        return cls(
            sap_system=os.getenv("SAP_SYSTEM", "DEV"),
            hana_host=os.getenv("HANA_HOST", "localhost"),
            hana_port=int(os.getenv("HANA_PORT", "30015")),
            hana_user=os.getenv("HANA_USER", "SYSTEM"),
            hana_password=os.getenv("HANA_PASSWORD", ""),
        )
