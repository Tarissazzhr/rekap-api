import pandas as pd
import calendar
from io import BytesIO

# ==========================
# Mapping hari Inggris ke Indonesia
# ==========================
day_map = {
    "Monday": "Senin",
    "Tuesday": "Selasa",
    "Wednesday": "Rabu",
    "Thursday": "Kamis",
    "Friday": "Jumat",
    "Saturday": "Sabtu",
    "Sunday": "Minggu"
}

# ==========================
# Fungsi utama (dipanggil API FastAPI)
# ==========================
def generate_rekap(absen_files, cuti_files):
    # ==========================
    # Load file absen
    # ==========================
    if absen_files:
        absen_list = [pd.read_excel(f) for f in absen_files]
        absen = pd.concat(absen_list, ignore_index=True)
    else:
        absen = pd.DataFrame(columns=["ID", "NIK", "User Name", "Department", 
                                      "Date", "First-In Time", "Last-Out Time"])

    # ==========================
    # Load file cuti
    # ==========================
    if cuti_files:
        cuti_list = [pd.read_excel(f) for f in cuti_files]
        cuti = pd.concat(cuti_list, ignore_index=True)
    else:
        cuti = pd.DataFrame(columns=["NIK", "Start Date", "End Date", "Reason Cuti"])

    # ==========================
    # Format tanggal & normalisasi NIK
    # ==========================
    if not absen.empty:
        absen["Date"] = pd.to_datetime(absen["Date"], dayfirst=True, errors="coerce")
        absen["NIK"] = absen["NIK"].astype(str).str.strip()

    if not cuti.empty:
        cuti["Start Date"] = pd.to_datetime(cuti["Start Date"], dayfirst=True, errors="coerce")
        cuti["End Date"] = pd.to_datetime(cuti["End Date"], dayfirst=True, errors="coerce")
        cuti["NIK"] = cuti["NIK"].astype(str).str.strip()

    # ==========================
    # Kalau tidak ada absen, return None
    # ==========================
    if absen.empty:
        return None

    # ==========================
    # Buat range tanggal
    # ==========================
    min_date = absen["Date"].min()
    max_date = absen["Date"].max()
    all_dates = pd.date_range(start=min_date, end=max_date)

    # ==========================
    # Data karyawan
    # ==========================
    employees = absen[["NIK", "User Name", "Department"]].drop_duplicates()
    employees["NIK"] = employees["NIK"].astype(str).str.strip()

    wide = employees.copy()
    wide.insert(0, "No", range(1, len(wide) + 1))

    # ==========================
    # MultiIndex header
    # ==========================
    tuples = [("No", "", ""), ("NIK", "", ""), ("User Name", "", ""), ("Department", "", "")]

    for d in all_dates:
        tgl = d.strftime("%d/%m/%Y")
        hari = day_map[d.strftime("%A")]
        tuples.extend([
            (tgl, hari, "In"),
            (tgl, hari, "Out"),
            (tgl, hari, "Reason")
        ])

    # Summary
    tuples.extend([
        ("Summary", "", "Jumlah Absen"),
        ("Summary", "", "Tidak Absen"),
        ("Summary", "", "Jumlah Cuti"),
        ("Summary", "", "Reason Cuti"),
    ])

    multi_index = pd.MultiIndex.from_tuples(tuples)
    result = pd.DataFrame(columns=multi_index)

    # Isi data dasar
    result[("No", "", "")] = wide["No"].values
    result[("NIK", "", "")] = wide["NIK"].values
    result[("User Name", "", "")] = wide["User Name"].values
    result[("Department", "", "")] = wide["Department"].values

    # ==========================
    # Isi absen In/Out
    # ==========================
    for _, row in absen.iterrows():
        nik = row["NIK"]
        d = row["Date"]
        tgl = d.strftime("%d/%m/%Y")
        hari = day_map[d.strftime("%A")]

        in_time = row["First-In Time"]
        out_time = row["Last-Out Time"]

        if str(in_time).startswith("00:00"):
            in_time = "-"
        if str(out_time).startswith("00:00"):
            out_time = "-"

        idx = result[("NIK", "", "")] == nik
        result.loc[idx, (tgl, hari, "In")] = in_time
        result.loc[idx, (tgl, hari, "Out")] = out_time

    # ==========================
    # Cuti
    # ==========================
    reason_dict = {}

    if not cuti.empty:
        for _, row in cuti.iterrows():
            nik = row["NIK"]
            reason = row["Reason Cuti"]

            reason_dict.setdefault(nik, set()).add(reason)

            mask = (all_dates >= row["Start Date"]) & (all_dates <= row["End Date"])

            for d in all_dates[mask]:
                tgl = d.strftime("%d/%m/%Y")
                hari = day_map[d.strftime("%A")]

                idx = result[("NIK", "", "")] == nik
                result.loc[idx, (tgl, hari, "In")] = "Cuti"
                result.loc[idx, (tgl, hari, "Out")] = "Cuti"

    # ==========================
    # Summary
    # ==========================
    jumlah_absen = []
    tidak_absen = []
    jumlah_cuti = []
    reason_list = []

    in_cols = [col for col in result.columns if col[2] == "In"]

    for _, row in result.iterrows():
        hadir = 0
        tidak_hadir = 0
        cuti_count = 0

        for tgl, hari, _ in in_cols:
            v = row[(tgl, hari, "In")]

            if v == "Cuti":
                cuti_count += 1
                continue

            dt = pd.to_datetime(tgl, dayfirst=True)
            if dt.weekday() in [5, 6]:
                continue

            if v in ["", "-", None] or pd.isna(v):
                tidak_hadir += 1
            else:
                hadir += 1

        nik = row[("NIK", "", "")]
        reason_text = ", ".join(reason_dict.get(nik, ["-"]))

        jumlah_absen.append(hadir)
        tidak_absen.append(tidak_hadir)
        jumlah_cuti.append(cuti_count)
        reason_list.append(reason_text)

    result[("Summary", "", "Jumlah Absen")] = jumlah_absen
    result[("Summary", "", "Tidak Absen")] = tidak_absen
    result[("Summary", "", "Jumlah Cuti")] = jumlah_cuti
    result[("Summary", "", "Reason Cuti")] = reason_list

    # ==========================
    # Output buffer Excel (untuk API)
    # ==========================
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result.to_excel(writer, sheet_name="Rekap", index=True)

    output.seek(0)
    return output
