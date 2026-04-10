"""Base agent interface."""

from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from typing import Any


class BaseAgent(ABC):
    """Common interface for all pipeline agents."""

    name: str = "BaseAgent"

    def __init__(self) -> None:
        self.logger = logging.getLogger(self.name)

    @abstractmethod
    def run(self, **kwargs: Any) -> Any:
        ...
