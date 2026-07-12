from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


def model_to_dict(model: Any) -> dict:
    if hasattr(model, "model_dump"):
        return model.model_dump()
    return model.dict()


class ProfilePayload(BaseModel):
    type: str = "管理"
    title: str = ""
    major: str = ""
    company: str = ""
    industry: str = ""
    design_type: str = ""
    design_object: str = ""
    context: str = ""


class GeneratePayload(BaseModel):
    type: str = "管理"
    profile: ProfilePayload = Field(default_factory=ProfilePayload)


class TaskStartResponse(BaseModel):
    status: str
    task_id: str


class TaskSnapshot(BaseModel):
    status: str = "missing"
    msg: str = "任务不存在"
    progress: int = 0
    files: list[str] = Field(default_factory=list)

