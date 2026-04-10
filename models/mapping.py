"""Structure mapping models for draft -> STO conversion."""

from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class ActionType(str, Enum):
    KEEP = "keep"
    RENAME = "rename"
    REORDER = "reorder"
    ADD = "add"
    MERGE = "merge"
    SPLIT = "split"
    REMOVE = "remove"


class SectionAction(BaseModel):
    action: ActionType
    source_number: Optional[str] = None
    source_title: Optional[str] = None
    target_number: str = ""
    target_title: str = ""
    notes: str = ""


class SectionMapping(BaseModel):
    draft_number: Optional[str] = None
    draft_title: Optional[str] = None
    sto_number: str = ""
    sto_title: str = ""
    action: ActionType = ActionType.KEEP
    content_instructions: str = ""


class StructureMapping(BaseModel):
    """Complete mapping plan from draft to STO structure."""

    mappings: list[SectionMapping] = Field(default_factory=list)
    actions: list[SectionAction] = Field(default_factory=list)
    sto_number: str = ""
    document_title: str = ""
    introduction_status: str = "Введен впервые"
    notes: str = ""
