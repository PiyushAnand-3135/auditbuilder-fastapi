from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

# Import your routers
from routes.iso9001 import router as iso9001_router
from routes.iso14001 import router as iso14001_router
from routes.iso45001 import router as iso45001_router
from routes.iso9001_14001 import router as iso9001_14001_router
from routes.iso9001_14001_45001 import router as iso9001_14001_45001_router

# Initialize FastAPI app
app = FastAPI(
    title="Accurate Report API",
    version="1.0.0",
    description="API for audit-builder reports"
)

# CORS Middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://report-frontend-dm1bya-519963-31-97-117-80.traefik.me"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Root route
@app.get("/")
def read_root():
    return {"message": "Welcome to Accurate Report API"}

# Include routers
app.include_router(iso9001_router, prefix="/iso9001", tags=["ISO 9001"])
app.include_router(iso14001_router, prefix="/iso14001", tags=["ISO 14001"])
app.include_router(iso45001_router, prefix="/iso45001", tags=["ISO 45001"])
app.include_router(iso9001_14001_router, prefix="/iso9001_14001", tags=["ISO 9001 + ISO 14001"])
app.include_router(iso9001_14001_45001_router, prefix="/iso9001_14001_45001", tags=["ISO 9001 + ISO 14001 + ISO 45001"])