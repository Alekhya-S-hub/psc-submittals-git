"""
Submittal Extraction API
FastAPI application for extracting submittals from construction spec books
"""
from fastapi import Request
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import Dict, List, Optional
import logging
import time
import tempfile
import os
from pathlib import Path
import pandas as pd
from io import BytesIO
import zipfile

from extractor import SubmittalExtractor

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="Submittal Extraction API",
    description="Extract submittal requirements from construction specification books",
    version="1.0.0",
    docs_url="/docs",
    redoc_url="/redoc"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/")
async def root():
    """
    Root endpoint - API information
    """
    return {
        "message": "Submittal Extraction API",
        "version": "1.0.0",
        "endpoints": {
            "health": "/health",
            "extract_excel": "/extract-submittals (returns ZIP with Excel files)",
            "extract_json": "/extract-submittals-json (returns JSON for n8n)",
            "docs": "/docs"
        }
    }


@app.get("/health")
async def health_check():
    """
    Health check endpoint
    Used by Azure and monitoring services
    """
    return {
        "status": "healthy",
        "service": "submittal-extraction-api",
        "version": "1.0.0",
        "timestamp": time.time()
    }


@app.post("/extract-submittals")
async def extract_submittals(file: UploadFile = File(...)):
    """
    Extract submittal requirements from PDF spec book
    Returns a ZIP file containing two Excel files

    Args:
        file: PDF file upload (multipart/form-data)

    Returns:
        ZIP file containing:
        - submittal_sections.xlsx (multiple sheets, one per section)
        - submittals_log.xlsx (from template)

    Raises:
        HTTPException: If file is invalid or extraction fails
    """
    start_time = time.time()
    temp_file_path = None

    try:
        # Validate file type
        if not file.filename.lower().endswith('.pdf'):
            raise HTTPException(
                status_code=400,
                detail="Invalid file type. Only PDF files are supported."
            )

        logger.info(f"Starting extraction for file: {file.filename}")

        # Save uploaded file to temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', mode='wb') as temp_file:
            temp_file_path = temp_file.name
            content = await file.read()

            # Log file info for debugging
            logger.info(f"Received file: {file.filename}, size: {len(content)} bytes")
            logger.info(f"First 4 bytes: {content[:4]}")  # Should be b'%PDF'

            temp_file.write(content)
            temp_file.flush()  # Ensure data is written to disk
            logger.info(f"File saved to temporary location: {temp_file_path}")

        try:
            with open(temp_file_path, 'rb') as f:
                header = f.read(4)
                if header != b'%PDF':
                    raise HTTPException(
                        status_code=400,
                        detail="Invalid PDF file - corrupted or not a PDF"
                    )
        except Exception as e:
            logger.error(f"PDF validation failed: {e}")
            raise HTTPException(
                status_code=400,
                detail=f"PDF validation failed: {str(e)}"
            )

        # Look for template
        template_path = Path(__file__).parent / "templates" / "SubmittalLog.xlsx"
        if not template_path.exists():
            # Try alternative locations
            template_path = Path("templates/SubmittalLog.xlsx")
            if not template_path.exists():
                template_path = None
                logger.warning("Template not found, will create basic structure")

        # Initialize extractor and process PDF
        extractor = SubmittalExtractor(temp_file_path)
        result = extractor.extract(template_path=str(template_path) if template_path else None)

        # Calculate extraction time
        extraction_time = time.time() - start_time

        # Get workbook objects
        sections_wb = result["sections"]  # openpyxl Workbook
        log_wb = result["log"]  # openpyxl Workbook

        logger.info(
            f"Extraction completed successfully: "
            f"{len(sections_wb.sheetnames)} section sheets, "
            f"Time: {extraction_time:.2f}s"
        )

        # Save workbooks to BytesIO
        sections_excel = BytesIO()
        log_excel = BytesIO()

        sections_wb.save(sections_excel)
        sections_excel.seek(0)

        log_wb.save(log_excel)
        log_excel.seek(0)

        # Create ZIP file containing both Excel files
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            # Add sections Excel
            zip_file.writestr('submittal_sections.xlsx', sections_excel.getvalue())
            # Add log Excel
            zip_file.writestr('submittals_log.xlsx', log_excel.getvalue())

        zip_buffer.seek(0)

        # Generate filename based on original PDF name
        pdf_name = Path(file.filename).stem
        zip_filename = f"{pdf_name}_extracted_submittals.zip"

        logger.info(f"Returning ZIP file: {zip_filename}")

        # Return ZIP file
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={
                "Content-Disposition": f"attachment; filename={zip_filename}"
            }
        )

    except HTTPException:
        raise

    except Exception as e:
        logger.error(f"Error during extraction: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Extraction failed: {str(e)}"
        )

    finally:
        # Clean up temporary file
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
                logger.info(f"Temporary file deleted: {temp_file_path}")
            except Exception as e:
                logger.warning(f"Failed to delete temporary file: {str(e)}")


@app.post("/extract-submittals-json")
async def extract_submittals_json(file: UploadFile = File(...)):
    """
    Extract submittal requirements and return JSON
    (For n8n or programmatic integration)

    Args:
        file: PDF file upload (multipart/form-data)

    Returns:
        JSON response with sections and log data
    """
    start_time = time.time()
    temp_file_path = None

    try:
        # Validate file type
        if not file.filename.lower().endswith('.pdf'):
            raise HTTPException(
                status_code=400,
                detail="Invalid file type. Only PDF files are supported."
            )

        logger.info(f"Starting extraction for file: {file.filename}")

        # Save uploaded file to temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
            temp_file_path = temp_file.name
            content = await file.read()
            temp_file.write(content)
            logger.info(f"File saved to temporary location: {temp_file_path}")

        # Initialize extractor and process PDF
        extractor = SubmittalExtractor(temp_file_path)
        result = extractor.extract()

        # Calculate extraction time
        extraction_time = time.time() - start_time

        logger.info(
            f"Extraction completed successfully: "
            f"{len(result['sections'])} sections, "
            f"{len(result['log'])} submittals"
        )

        # Prepare response
        response = {
            "success": True,
            "sections": result["sections"],
            "log": result["log"],
            "metadata": {
                "filename": file.filename,
                "total_sections": len(result["sections"]),
                "total_submittals": len(result["log"]),
                "extraction_time": round(extraction_time, 2),
                "timestamp": time.time()
            }
        }

        return JSONResponse(content=response)

    except HTTPException:
        raise

    except Exception as e:
        logger.error(f"Error during extraction: {str(e)}", exc_info=True)
        raise HTTPException(
            status_code=500,
            detail=f"Extraction failed: {str(e)}"
        )

    finally:
        # Clean up temporary file
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
                logger.info(f"Temporary file deleted: {temp_file_path}")
            except Exception as e:
                logger.warning(f"Failed to delete temporary file: {str(e)}")



@app.post("/extract-submittals-base64")
async def extract_submittals_base64(request: Request):
    """
    Base64 endpoint with raw request parsing - for debugging
    """
    start_time = time.time()
    temp_file_path = None

    try:
        # Parse JSON manually
        body = await request.json()

        filename = body.get("filename")
        file_content = body.get("file_content")

        logger.info(f"Received filename: {filename}")
        logger.info(f"Received file_content length: {len(file_content) if file_content else 0}")

        if not file_content:
            raise HTTPException(400, "file_content is required")

        if not filename:
            filename = "document.pdf"

        if not filename.lower().endswith('.pdf'):
            raise HTTPException(400, "Invalid file type")

        # Decode base64
        import base64
        file_bytes = base64.b64decode(file_content)
        logger.info(f"Decoded file size: {len(file_bytes)} bytes")

        # Validate PDF header
        if file_bytes[:4] != b'%PDF':
            raise HTTPException(400, f"Invalid PDF - first bytes: {file_bytes[:4]}")

        # Save to temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf', mode='wb') as temp_file:
            temp_file_path = temp_file.name
            temp_file.write(file_bytes)
            temp_file.flush()
            logger.info(f"File saved to: {temp_file_path}")

        # Look for template
        template_path = Path(__file__).parent / "templates" / "SubmittalLog.xlsx"
        if not template_path.exists():
            template_path = Path("templates/SubmittalLog.xlsx")
            if not template_path.exists():
                template_path = None

        # Extract
        extractor = SubmittalExtractor(temp_file_path)
        result = extractor.extract(template_path=str(template_path) if template_path else None)

        extraction_time = time.time() - start_time
        sections_wb = result["sections"]
        log_wb = result["log"]

        logger.info(f"Extraction completed: {len(sections_wb.sheetnames)} sheets, {extraction_time:.2f}s")

        # Create ZIP
        sections_excel = BytesIO()
        log_excel = BytesIO()

        sections_wb.save(sections_excel)
        log_wb.save(log_excel)

        sections_excel.seek(0)
        log_excel.seek(0)

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            zip_file.writestr('submittal_sections.xlsx', sections_excel.getvalue())
            zip_file.writestr('submittals_log.xlsx', log_excel.getvalue())

        zip_buffer.seek(0)

        pdf_name = Path(filename).stem
        zip_filename = f"{pdf_name}_extracted_submittals.zip"

        logger.info(f"Returning ZIP: {zip_filename}")

        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": f"attachment; filename={zip_filename}"}
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error: {str(e)}", exc_info=True)
        raise HTTPException(500, f"Extraction failed: {str(e)}")
    finally:
        if temp_file_path and os.path.exists(temp_file_path):
            try:
                os.unlink(temp_file_path)
            except Exception as e:
                logger.warning(f"Cleanup failed: {str(e)}")
@app.get("/test")
async def test_endpoint():
    """
    Simple test endpoint to verify API is responding
    """
    return {
        "status": "API is working",
        "message": "Test successful",
        "timestamp": time.time()
    }


# Error handlers
@app.exception_handler(404)
async def not_found_handler(request, exc):
    return JSONResponse(
        status_code=404,
        content={"error": "Not found", "detail": "The requested resource was not found"}
    )


@app.exception_handler(500)
async def internal_error_handler(request, exc):
    logger.error(f"Internal server error: {str(exc)}", exc_info=True)
    return JSONResponse(
        status_code=500,
        content={"error": "Internal server error", "detail": "An unexpected error occurred"}
    )


# Startup event
@app.on_event("startup")
async def startup_event():
    logger.info("Submittal Extraction API starting up...")
    logger.info("API is ready to accept requests")


# Shutdown event
@app.on_event("shutdown")
async def shutdown_event():
    logger.info("Submittal Extraction API shutting down...")


if __name__ == "__main__":
    import uvicorn

    # Get port from environment variable (Azure uses PORT)
    port = int(os.getenv("PORT", 8000))

    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=port,
        reload=True,  # Set to False in production
        log_level="info"
    )