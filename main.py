import httpx
import tempfile, os
import signal
import sys
import asyncio
from typing import Dict
from fastapi.responses import JSONResponse
from helper_functions import run_matrix_pipeline
from fastapi.middleware.cors import CORSMiddleware
from fastapi import FastAPI, File, UploadFile, Form, status, HTTPException
from contextlib import asynccontextmanager
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# --- Configuration ---
# Dynamically get backend URL from environment variable, default to localhost for development
backend_url = os.getenv("BACKEND_URL") 

SERVICE_API_KEY = os.getenv("SERVICE_API_KEY")

# Graceful shutdown handling
def signal_handler(signum, frame):
    """Handle CTRL+C and other termination signals gracefully."""
    print(f"\n🛑 Received signal {signum}. Initiating graceful shutdown...")
    print("👋 FastAPI server shutting down gracefully.")
    sys.exit(0)

# Register signal handlers for graceful shutdown
signal.signal(signal.SIGINT, signal_handler)   # CTRL+C
signal.signal(signal.SIGTERM, signal_handler)  # Termination signal

@asynccontextmanager
async def lifespan(app: FastAPI):
    """
    Handles startup and shutdown events for the FastAPI application.
    Fetches necessary data on startup using service-to-service authentication.
    """
    global faculty_abbreviations, subject_abbreviations
    try:
        # Validate that the SERVICE_API_KEY is configured
        if not SERVICE_API_KEY:
            raise ValueError("SERVICE_API_KEY environment variable is not set")
                
    except httpx.HTTPStatusError as e:
        print(f"❌ HTTP Error during startup data fetch: {e}")
        print(f"Request URL: {e.request.url}")
        print(f"Response Status: {e.response.status_code}")
        print(f"Response Body: {e.response.text}")
        raise RuntimeError(f"Failed to fetch essential startup data: {e.response.text}") from e
    except Exception as e:
        print(f"❌ An unexpected error occurred during startup data fetch: {e}")
        raise RuntimeError(f"Failed to fetch essential startup data: {e}") from e
    yield
    # Clean up on shutdown if necessary
    print("👋 Graceful Exit")
# Initialize FastAPI app with the lifespan manager
app = FastAPI(lifespan=lifespan)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {
        "status": "healthy",
        "message": "FastAPI Server is Live!",
        "service": "Faculty Matrix Processing",
        "version": "2.0.0"
    }

@app.post("/api/faculty-matrix", response_model=Dict, status_code=status.HTTP_200_OK)
async def faculty_matrix(
    facultyMatrix: UploadFile = File(...),
    deptAbbreviation: str = Form(...)
):
    temp_file_path = None
    try:
        # Validate file type
        if not facultyMatrix.filename:
            raise HTTPException(
                status_code=400, 
                detail="No file uploaded"
            )
        
        # Check file extension
        suffix = os.path.splitext(facultyMatrix.filename)[-1].lower()
        if suffix not in ['.xlsx', '.xls']:
            raise HTTPException(
                status_code=400, 
                detail="Invalid file type. Only Excel files (.xlsx, .xls) are supported"
            )
        
        # Validate department abbreviation
        if not deptAbbreviation or not deptAbbreviation.strip():
            raise HTTPException(
                status_code=400, 
                detail="Department abbreviation is required"
            )
        
        # Create temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
            try:
                content = await facultyMatrix.read()
                if not content:
                    raise HTTPException(
                        status_code=400, 
                        detail="Uploaded file is empty"
                    )
                temp_file.write(content)
                temp_file_path = temp_file.name
            except Exception as e:
                raise HTTPException(
                    status_code=400, 
                    detail=f"Error reading uploaded file: {str(e)}"
                )

        # Process the matrix
        try:
            results = run_matrix_pipeline(
                matrix_file_path=temp_file_path,
                department=deptAbbreviation.strip(),
                college="LDRP-ITR"
            )
        except Exception as e:
            print(f"❌ Matrix processing error: {str(e)}")
            raise HTTPException(
                status_code=500, 
                detail=f"Error processing timetable matrix: {str(e)}"
            )

        if not results:
            print("❌ No results found in the processed matrix")
            raise HTTPException(
                status_code=404, 
                detail="No timetable data found in the uploaded matrix. Please check the file format and content."
            )

        print(f"✅ Successfully processed matrix for department: {deptAbbreviation}")
        return JSONResponse(content=results)

    except HTTPException:
        raise
    except Exception as e:
        print(f"❌ Unexpected error in faculty_matrix endpoint: {str(e)}")
        raise HTTPException(
            status_code=500, 
            detail=f"An unexpected error occurred: {str(e)}"
        )
    finally:
        # Cleanup temporary file
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
                print(f"🗑️ Cleaned up temporary file: {temp_file_path}")
            except Exception as e:
                print(f"⚠️ Warning: Could not delete temporary file {temp_file_path}: {str(e)}")