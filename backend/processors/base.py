from abc import ABC, abstractmethod
from typing import Dict, Any

class BaseProcessor(ABC):
    @abstractmethod
    def get_category(self) -> str:
        pass

    @abstractmethod
    def process_sheet(self, ws, filename: str, master_equip: Dict[str, str], **kwargs) -> Dict[str, Any]:
        pass
