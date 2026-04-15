from dataclasses import dataclass
from pathlib import Path


@dataclass
class AppConfig:
    input_dir: Path
    output_dir: Path
    supported_extensions: tuple[str, ...] = (".txt", ".csv", ".xlsx", ".xlsm")

    @classmethod
    def default(cls, base_dir: Path | None = None) -> "AppConfig":
        root_dir = base_dir or Path.cwd()
        data_dir = root_dir / "data"
        return cls(
            input_dir=data_dir / "input",
            output_dir=data_dir / "output",
        )
