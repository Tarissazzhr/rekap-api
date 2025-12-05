# backend/main.py
from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
import pandas as pd
import io

app = FastAPI()

# CORS supaya bisa diakses dari Next.js
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # atau masukkan URL frontend mu
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mapping hari Inggris → Indonesia
day_map = {
    "Monday": "Senin",
    "Tuesday": "Selasa",
    "Wednesday": "Rabu",
    "Thursday": "Kamis",
    "Friday": "Jumat",
    "Saturday": "Sabtu",
    "Sunday": "Minggu"
}

@app.post("/rekap")
async def rekap_absen(
    absen_file: UploadFile = File(...),
    cuti_file: UploadFile = File(...)
):
    # Baca file Excel
    absen = pd.read_excel(absen_file.file)
    cuti = pd.read_excel(cuti_file.file)

    # Normalisasi NIK & Date
    if not absen.empty:
        absen["Date"] = pd.to_datetime(absen["Date"], dayfirst=True, errors="coerce")
        absen["NIK"] = absen["NIK"].astype(str).str.strip()
    if not cuti.empty:
        cuti["Start Date"] = pd.to_datetime(cuti["Start Date"], dayfirst=True, errors="coerce")
        cuti["End Date"] = pd.to_datetime(cuti["End Date"], dayfirst=True, errors="coerce")
        cuti["NIK"] = cuti["NIK"].astype(str).str.strip()

    if absen.empty:
        return {"detail": "Tidak ada data absen untuk direkap."}

    # Proses rekap
    min_date = absen["Date"].min()
    max_date = absen["Date"].max()
    all_dates = pd.date_range(start=min_date, end=max_date)

    employees = absen[["NIK", "User Name", "Department"]].drop_duplicates()
    employees["NIK"] = employees["NIK"].astype(str).str.strip()
    wide = employees.copy()
    wide.insert(0, "No", range(1, len(wide)+1))

    # MultiIndex columns
    tuples = [("No", "", ""), ("NIK", "", ""), ("User Name", "", ""), ("Department", "", "")]
    for d in all_dates:
        tanggal = d.strftime("%d/%m/%Y")
        hari = day_map[d.strftime("%A")]
        tuples.append((tanggal, hari, "In"))
        tuples.append((tanggal, hari, "Out"))
        tuples.append((tanggal, hari, "Reason"))
    tuples.append(("Summary", "", "Jumlah Absen"))
    tuples.append(("Summary", "", "Tidak Absen"))
    tuples.append(("Summary", "", "Jumlah Cuti"))
    tuples.append(("Summary", "", "Reason Cuti"))
    multi_index = pd.MultiIndex.from_tuples(tuples)

    # Buat dataframe hasil
    result = pd.DataFrame(columns=multi_index)
    result[("No", "", "")] = wide["No"].values
    result[("NIK", "", "")] = wide["NIK"].values
    result[("User Name", "", "")] = wide["User Name"].values
    result[("Department", "", "")] = wide["Department"].values

    # Isi absen
    for _, row in absen.iterrows():
        emp_id = row["NIK"]
        d = row["Date"]
        tanggal = d.strftime("%d/%m/%Y")
        hari = day_map[d.strftime("%A")]

        in_time = row["First-In Time"]
        out_time = row["Last-Out Time"]

        if str(in_time).startswith("00:00"): in_time = "-"
        if str(out_time).startswith("00:00"): out_time = "-"

        result.loc[result[("NIK", "", "")] == emp_id, (tanggal, hari, "In")] = in_time
        result.loc[result[("NIK", "", "")] == emp_id, (tanggal, hari, "Out")] = out_time

    # Cuti
    reason_dict = {}
    if not cuti.empty:
        for _, row in cuti.iterrows():
            emp_id = row["NIK"]
            reason = row["Reason Cuti"]
            reason_dict.setdefault(emp_id, set()).add(reason)

            mask = (all_dates >= row["Start Date"]) & (all_dates <= row["End Date"])
            for d in all_dates[mask]:
                tanggal = d.strftime("%d/%m/%Y")
                hari = day_map[d.strftime("%A")]
                idx = result[("NIK", "", "")] == emp_id
                result.loc[idx, (tanggal, hari, "In")] = "Cuti"
                result.loc[idx, (tanggal, hari, "Out")] = "Cuti"

    # Summary
    jumlah_absen, tidak_absen, jumlah_cuti, reason_list = [], [], [], []
    in_cols = [col for col in result.columns if col[2]=="In"]

    for idx, row in result.iterrows():
        hadir = tidak_hadir = cuti_count = 0
        for tanggal, hari, tipe in in_cols:
            v = row[(tanggal, hari, "In")]
            if v=="Cuti": cuti_count += 1; continue
            dt = pd.to_datetime(tanggal, dayfirst=True)
            if dt.weekday() in [5,6]: continue
            if v in ["", "-", None] or pd.isna(v): tidak_hadir += 1
            else: hadir += 1
        emp_id = row[("NIK", "", "")]
        reason_text = ", ".join(reason_dict.get(emp_id, ["-"])) if not cuti.empty else "-"
        jumlah_absen.append(hadir)
        tidak_absen.append(tidak_hadir)
        jumlah_cuti.append(cuti_count)
        reason_list.append(reason_text)

    result[("Summary", "", "Jumlah Absen")] = jumlah_absen
    result[("Summary", "", "Tidak Absen")] = tidak_absen
    result[("Summary", "", "Jumlah Cuti")] = jumlah_cuti
    result[("Summary", "", "Reason Cuti")] = reason_list

    # Simpan Excel di memory
    output = io.BytesIO()
    result.to_excel(output, index=True, sheet_name="Rekap")
    output.seek(0)

    return StreamingResponse(output,
                             media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             headers={"Content-Disposition": "attachment; filename=rekap.xlsx"})
