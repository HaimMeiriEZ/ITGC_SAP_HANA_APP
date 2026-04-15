import unittest

from src.config import AppConfig
from src.services.hana_service import HanaService


class TestSmoke(unittest.TestCase):
    def test_config_defaults(self) -> None:
        config = AppConfig.from_env()
        self.assertEqual(config.sap_system, "DEV")

    def test_hana_service_healthcheck(self) -> None:
        service = HanaService("localhost", 30015, "SYSTEM")
        result = service.healthcheck()
        self.assertEqual(result["status"], "ready")


if __name__ == "__main__":
    unittest.main()
