from fastapi import FastAPI, UploadFile, Form, Response
from fastapi.middleware.cors import CORSMiddleware
from pptx import Presentation
import io
import json
import logging
import time
import os
from datetime import datetime

app = FastAPI()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
)
logger = logging.getLogger("pptx-generator-api")

# Create output directory for local cross-reference files
OUTPUT_DIR = "./generated_decks"
os.makedirs(OUTPUT_DIR, exist_ok=True)
logger.info("Output directory for generated files: %s", os.path.abspath(OUTPUT_DIR))

# 1. CORS Configuration: Allow your SPFx web part to communicate with this server
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"], # For production, replace "*" with "https://yourtenant.sharepoint.com"
    allow_methods=["POST"],
    allow_headers=["*"],
)

def delete_slide(prs, index):
    """Helper function to remove a slide by its index."""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[index])

@app.post("/generate-document")
async def generate_document(
    file: UploadFile, 
    data: str = Form(...) # The JSON string sent from SPFx containing variables and instructions
):
    started_at = time.perf_counter()
    logger.info("Request received at /generate-document")
    logger.info("Incoming file name: %s", file.filename)

    # 2. Parse the incoming JSON data
    instructions = json.loads(data)
    variables = instructions.get("variables", {})
    slides_to_remove = instructions.get("deleteSlides", [])
    metadata = instructions.get("metadata", {})

    logger.info("Variables count: %s", len(variables))
    logger.info("Slides requested to delete: %s", slides_to_remove)
    logger.info("Metadata: %s", metadata)

    # 3. Read the incoming file (Blob) into server memory
    file_bytes = await file.read()
    logger.info("Input file size (bytes): %s", len(file_bytes))
    input_stream = io.BytesIO(file_bytes)
    prs = Presentation(input_stream)
    logger.info("Loaded presentation with %s slides", len(prs.slides))

    # 4. Delete specified slides (Must be done in reverse order to avoid index shifting)
    deleted_count = 0
    for index in sorted(slides_to_remove, reverse=True):
        if index < len(prs.slides):
            delete_slide(prs, index)
            deleted_count += 1
            logger.info("Deleted slide at index: %s", index)
        else:
            logger.warning("Skipped invalid slide index: %s", index)
    logger.info("Total deleted slides: %s", deleted_count)

    # 5. Search and Replace text in shapes and tables
    replacement_count = 0
    for slide in prs.slides:
        for shape in slide.shapes:
            # Handle standard text boxes
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in variables.items():
                            if key in run.text:
                                run.text = run.text.replace(key, str(value))
                                replacement_count += 1
            
            # Handle tables
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                for key, value in variables.items():
                                    if key in run.text:
                                        run.text = run.text.replace(key, str(value))
                                        replacement_count += 1
    logger.info("Total text replacements applied: %s", replacement_count)

    # 6. Save the modified presentation to a new memory buffer
    output_stream = io.BytesIO()
    prs.save(output_stream)
    logger.info("Presentation saved to memory buffer")
    
    # Reset the buffer's "cursor" to the beginning before reading it out
    output_stream.seek(0)
    output_bytes = output_stream.getvalue()

    # 6.5 Save to local file for cross-referencing
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    search_keyword = metadata.get("currentSearchKeyword", "unknown").replace(" ", "_")
    local_filename = f"nijobair_{search_keyword}_{timestamp}.pptx"
    local_filepath = os.path.join(OUTPUT_DIR, local_filename)
    
    try:
        with open(local_filepath, "wb") as f:
            f.write(output_bytes)
        logger.info("Local copy saved: %s", local_filepath)
    except Exception as e:
        logger.error("Failed to save local copy: %s", str(e))

    elapsed_ms = int((time.perf_counter() - started_at) * 1000)
    logger.info("Returning modified presentation. Output size: %s bytes", len(output_bytes))
    logger.info("Request processing completed in %s ms", elapsed_ms)

    # 7. Return the raw binary stream (Blob) directly back to SPFx
    return Response(
        content=output_bytes,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=nijobair.pptx"},
    )

@app.get("/")
async def root():
    return {"message": "Welcome to the PowerPoint Generator API"}