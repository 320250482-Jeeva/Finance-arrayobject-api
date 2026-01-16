from fastapi import FastAPI
from routers import array_router, ppt_email_router

app = FastAPI()

app.include_router(array_router)
app.include_router(ppt_email_router)
