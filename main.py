from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from typing import List, Optional
from rekap_core import generate_rekap
import io

app = FastAPI()

# ================================
# CORS (untuk Next.js atau domain lainnya)
# ================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # boleh semua domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "Absensi API Ready!"}


# ================================
# ENDPOINT: /rekap
# ================================
@app.post("/rekap")
async def rekap_absen(
    absen_files: Optional[List[UploadFile]] = File(None),
    cuti_files: Optional[List[UploadFile]] = File(None)
):
    """
    Upload:
        absen_files[]: Excel file absen
        cuti_files[] : Excel file cuti (opsional)

    Return:
        File Excel Rekap (StreamingResponse)
    """

    # Convert file → byte stream untuk core engine
    absen_streams = [io.BytesIO(await f.read()) for f in absen_files] if absen_files else []
    cuti_streams = [io.BytesIO(await f.read()) for f in cuti_files] if cuti_files else []

    # Generate Excel rekap
    output = generate_rekap(absen_streams, cuti_streams)

    if output is None:
        return JSONResponse(
            status_code=400,
            content={"error": "Tidak ada data absen valid di file yang diupload."}
        )

    # Return file Excel
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=rekap_absen.xlsx"
        }
    )
