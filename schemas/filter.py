# schemas/filter.py
from pydantic import BaseModel, Field, field_validator
from typing import List, Optional

class FilterCondition(BaseModel):
    field: str
    values: List[str]

    @field_validator('field')
    @classmethod
    def field_not_empty(cls, v):
        if not v or not v.strip():
            raise ValueError('字段名不能为空')
        return v.strip()

    @field_validator('values')
    @classmethod
    def values_not_empty(cls, v):
        if not v:
            raise ValueError('值列表不能为空')
        return [val.strip() for val in v if val.strip()]

class FilterParams(BaseModel):
    files: List[str] = Field(..., min_length=1, description="文件路径列表")
    conditions: List[FilterCondition] = Field(default_factory=list)
    fuzzy_field: Optional[str] = None
    fuzzy_value: Optional[str] = None
    action: str = Field(..., description="操作类型: get_values 或 generate_excel")
    target_field: Optional[str] = None

    @field_validator('action')
    @classmethod
    def validate_action(cls, v):
        if v not in ('get_values', 'generate_excel'):
            raise ValueError('action 必须是 get_values 或 generate_excel')
        return v

    @field_validator('target_field')
    @classmethod
    def validate_target_field(cls, v, info):
        if info.data.get('action') == 'get_values' and not v:
            raise ValueError('获取值时必须指定 target_field')
        return v