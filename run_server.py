import uvicorn
from services.api import app

if __name__ == "__main__":
    uvicorn.run(
        "services.api:app",
        host="0.0.0.0",
        port=8000,
        reload=True
    )
