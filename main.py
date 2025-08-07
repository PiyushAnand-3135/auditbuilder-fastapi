from fastapi import FastAPI, Response, Request
from fastapi.middleware.cors import CORSMiddleware
from routes.iso9001 import router as iso9001_router
from routes.iso14001 import router as iso14001_router
from routes.iso9001_14001 import router as iso9001_14001_router
from routes.iso9001_14001_45001 import router as iso9001_14001_45001_router
from routes.iso45001 import router as iso45001_router


app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(iso9001_router, prefix="/iso9001")
app.include_router(iso14001_router, prefix="/iso14001")
app.include_router(iso45001_router, prefix="/iso45001")
app.include_router(iso9001_14001_router, prefix="/iso9001_14001")
app.include_router(iso9001_14001_45001_router, prefix="/iso9001_14001_45001")

