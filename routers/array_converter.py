from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse

router = APIRouter(tags=["Array Converter"])

@router.get("/")
def root():
    return {"message": "Welcome to Array Converter API!"}

@router.post("/convert")
async def convert_arrays(request: Request):
    body = await request.json()
    header = body.get("header")
    data = body.get("data")

    if not header or not data:
        return JSONResponse(
            content={"error": "Both 'header' and 'data' are required."},
            status_code=400
        )

    return [dict(zip(header, row)) for row in data]
