from fastapi import FastAPI, UploadFile
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from rekap import generate_rekap
import io

app = FastAPI()

# ================================
# CORS (untuk Next.js di Vercel)
# ================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "Absensi API Ready!"}

# ================================
# ENDPOINT /rekap — menerima 2 file
# ================================
@app.post("/rekap")
async def rekap(absensi: UploadFile, cuti: UploadFile):

    # ====== Baca file absensi & cuti ======
    absensi_bytes = io.BytesIO(await absensi.read())
    cuti_bytes = io.BytesIO(await cuti.read())

    # convert ke list stream karena engine kamu generate_rekap expects list
    absen_streams = [absensi_bytes]
    cuti_streams = [cuti_bytes]

    # ====== Panggil engine rekap ======
    output_stream = generate_rekap(absen_streams, cuti_streams)

    if output_stream is None:
        return JSONResponse(
            status_code=400,
            content={"error": "File tidak valid atau tidak berisi data absen."}
        )

    # Pastikan pointer file di awal
    output_stream.seek(0)

    # ====== RETURN FILE EXCEL ======
    return StreamingResponse(
        output_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": "attachment; filename=rekap_absen.xlsx"
        }
    )
