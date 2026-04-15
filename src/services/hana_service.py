class HanaService:
    def __init__(self, host: str, port: int, user: str) -> None:
        self.host = host
        self.port = port
        self.user = user

    def healthcheck(self) -> dict:
        return {
            "status": "ready",
            "host": self.host,
            "port": self.port,
            "user": self.user,
        }
