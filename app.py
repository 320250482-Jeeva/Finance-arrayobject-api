from fastapi import FastAPI, Request
from fastapi.responses import JSONResponse

app = FastAPI(title="Array Converter API")

@app.get("/")
def root():
    return {"message": "Welcome to Array Converter API!"}

@app.post("/convert")
async def convert_arrays(request: Request):
    """
    Converts:
      header = ["name", "age"]
      data = [["John", 25], ["Emma", 30]]
    To:
      [{"name": "John", "age": 25}, {"name": "Emma", "age": 30}]
    """
    body = await request.json()
    header = body.get("header")
    data = body.get("data")

    # Validation
    if not header or not data:
        return JSONResponse(
            content={"error": "Both 'header' and 'data' are required."},
            status_code=400
        )

    # Convert
    result = [dict(zip(header, row)) for row in data]
    return result
