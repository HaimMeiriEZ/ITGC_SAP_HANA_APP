from src.config import AppConfig


def run() -> None:
    config = AppConfig.from_env()
    print("ITGC SAP HANA app initialized")
    print(f"SAP system: {config.sap_system}")
    print(f"HANA host: {config.hana_host}")


if __name__ == "__main__":
    run()
