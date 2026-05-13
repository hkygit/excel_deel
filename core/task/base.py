# core/task/base.py
from abc import ABC, abstractmethod
from typing import Any

class Task(ABC):
    @abstractmethod
    def execute(self) -> Any:
        pass
