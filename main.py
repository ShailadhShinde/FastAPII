from fastapi import FastAPI, UploadFile, File, HTTPException, Query, Path
from fastapi.responses import JSONResponse
from starlette.status import HTTP_201_CREATED
from urllib.parse import unquote
import re

from schemas import TableListResponse, TableDetailsResponse, RowSumResponse
from utils import process_excel_file

app = FastAPI(title="Excel Processor API")

db = {}

def clean_parameter(param: str) -> str:
    """Clean and decode URL parameters to handle special characters"""
    try:
        decoded = unquote(param)
        if decoded != param:
            return decoded.strip()
    except:
        pass
    
    return param.strip()

def find_table_name(requested_name: str) -> str:
    cleaned_name = clean_parameter(requested_name)
    
    # Exact match first
    if cleaned_name in db:
        return cleaned_name
    
    # Case-insensitive match
    for table_name in db.keys():
        if table_name.lower() == cleaned_name.lower():
            return table_name
    
    # Partial match (contains)
    for table_name in db.keys():
        if cleaned_name.lower() in table_name.lower() or table_name.lower() in cleaned_name.lower():
            return table_name
    
    # If no match found, raise error with suggestions
    available_tables = list(db.keys())
    raise HTTPException(404, f"Table '{cleaned_name}' not found. Available tables: {available_tables}")

def find_row_name(table_data: list, requested_row_name: str) -> int:
    """Find row index by name, handling special characters and partial matches"""
    cleaned_row_name = clean_parameter(requested_row_name)
    
    # First pass: exact match
    for i, row in enumerate(table_data):
        if row and len(row) > 0:
            first_cell = str(row[0]).strip()
            if first_cell == cleaned_row_name:
                return i
    
    # Second pass: case-insensitive match
    for i, row in enumerate(table_data):
        if row and len(row) > 0:
            first_cell = str(row[0]).strip()
            if first_cell.lower() == cleaned_row_name.lower():
                return i
    
    # Third pass: partial match
    for i, row in enumerate(table_data):
        if row and len(row) > 0:
            first_cell = str(row[0]).strip()
            if cleaned_row_name.lower() in first_cell.lower() or first_cell.lower() in cleaned_row_name.lower():
                return i
    
    # If no match found, raise error with available row names
    available_rows = []
    for row in table_data:
        if row and len(row) > 0 and str(row[0]).strip():
            available_rows.append(str(row[0]).strip())
    
    raise HTTPException(404, f"Row '{cleaned_row_name}' not found. Available rows: {available_rows[:10]}")  # Limit to first 10

def ensure_uploaded():
    """Ensure an Excel file has been uploaded"""
    if not db:
        raise HTTPException(400, "No Excel file uploaded yet")

@app.post(
    "/upload_excel",
    response_model=TableListResponse,
    status_code=HTTP_201_CREATED
)
async def upload_excel(file: UploadFile = File(...)):
    """Upload and process an Excel file"""
    fname = file.filename.lower()
    if not (fname.endswith(".xls") or fname.endswith(".xlsx")):
        raise HTTPException(400, "Only .xls/.xlsx files allowed")
    
    data = await file.read()
    try:
        tables = process_excel_file(data)
    except Exception as e:
        raise HTTPException(500, f"Error processing Excel: {e}")
    
    db.clear()
    db.update(tables)
    return TableListResponse(tables=list(db.keys()))

@app.get("/list_tables", response_model=TableListResponse)
def list_tables():
    """List all available tables from the uploaded Excel file"""
    ensure_uploaded()
    return TableListResponse(tables=list(db.keys()))

@app.get("/debug_table")
def debug_table(table_name: str = Query(..., min_length=1)):
    """Debug endpoint to see raw table data"""
    ensure_uploaded()
    
    # Use robust table name finding
    actual_table_name = find_table_name(table_name)
    table_data = db[actual_table_name]
    
    return {
        "requested_name": table_name,
        "actual_table_name": actual_table_name,
        "num_rows": len(table_data),
        "num_cols": len(table_data[0]) if table_data else 0,
        "first_5_rows": table_data[:5],
        "all_first_cells": [row[0] if row else "EMPTY_ROW" for row in table_data]
    }

@app.get("/get_table_details", response_model=TableDetailsResponse)
def get_table_details(table_name: str = Query(..., min_length=1)):
    """Get details of a specific table including all row names"""
    ensure_uploaded()
    
    # Use robust table name finding
    actual_table_name = find_table_name(table_name)
    table_data = db[actual_table_name]
    
    if not table_data:
        raise HTTPException(400, f"Table '{actual_table_name}' has no data")
    
    # More robust row name extraction
    row_names = []
    for i, row in enumerate(table_data):
        if not row:  # Empty row
            row_names.append(f"<empty_row_{i}>")
        elif not row[0] or row[0].strip() == "":  # Empty first cell
            row_names.append(f"<empty_cell_{i}>")
        else:
            row_names.append(str(row[0]).strip())
    
    return TableDetailsResponse(table_name=actual_table_name, row_names=row_names)

@app.get("/row_sum", response_model=RowSumResponse)
def row_sum(
    table_name: str = Query(..., min_length=1),
    row_name: str = Query(..., min_length=1)
):
    """Calculate the sum of numeric values in a specific row"""
    ensure_uploaded()
    
    # Use robust table and row name finding
    actual_table_name = find_table_name(table_name)
    table_data = db[actual_table_name]
    
    row_idx = find_row_name(table_data, row_name)
    actual_row_name = str(table_data[row_idx][0]).strip()
    
    total = 0.0
    for cell in table_data[row_idx][1:]:  # Skip first column (row name)
        try:
            total += float(cell)
        except:
            pass  # Skip non-numeric values
    
    return RowSumResponse(
        table_name=actual_table_name, 
        row_name=actual_row_name, 
        sum=round(total, 4)
    )

@app.exception_handler(HTTPException)
def http_error_handler(request, exc: HTTPException):
    """Global HTTP exception handler"""
    return JSONResponse(
        status_code=exc.status_code, 
        content={"error": exc.detail}
    )