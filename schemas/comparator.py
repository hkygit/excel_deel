# schemas/comparator.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class CompareFullParams(BaseModel):
    key_fields: List[str] = Field(..., min_length=1, description="主键字段列表")
    compare_fields: List[str] = Field(..., min_length=1, description="对比字段列表")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('key_fields', 'compare_fields', mode='before')
    @classmethod
    def check_not_empty(cls, v):
        if not v:
            raise ValueError('字段列表不能为空')
        return [item.strip() for item in v if item and item.strip()]

class CompareFieldMatchParams(BaseModel):
    match_field: str = Field(..., description="匹配字段名")
    sheet1: Optional[str] = None
    sheet2: Optional[str] = None

    @field_validator('match_field')
    @classmethod
    def not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('匹配字段名不能为空')
        return v.strip()