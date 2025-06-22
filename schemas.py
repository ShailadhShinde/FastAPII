from pydantic import BaseModel

class TableListResponse(BaseModel):
    tables: list[str]

class TableDetailsResponse(BaseModel):
    table_name: str
    row_names: list[str]

class RowSumResponse(BaseModel):
    table_name: str
    row_name: str
    sum: float
