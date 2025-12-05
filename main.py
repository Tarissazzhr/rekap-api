from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from rekap import generate_rekap
import io

app = FastAPI()

# CORS supaya Next.js bisa fetch API ini
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

@app.post("/rekap")
async def rekap_endpoint(
    absensi: UploadFile = File(...),
    cuti: UploadFile = File(...)
):
    try:
        absen_stream = io.BytesIO(await absensi.read())
        cuti_stream = io.BytesIO(await cuti.read())

        # Call core logic
        output_stream = generate_rekap([absen_stream], [cuti_stream])

        if output_stream is None:
            return JSONResponse(
                status_code=400,
                content={"error": "Gagal membuat rekap. File mungkin tidak valid."}
            )

        return StreamingResponse(
            output_stream,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition": "attachment; filename=rekap_absen.xlsx"
            }
        )
    except Exception as e:
        return JSONResponse(
            status_code=500,
            content={"error": str(e)}
        )
