"""
THC Data Merger
================
Streamlit app untuk menggabungkan (merge) beberapa file transaksi bulanan THC
(hasil ekspor MDIS: .xlsb / .xls / .xlsx / .csv) menjadi satu file rapi dengan
header standar:

    VOUCHER NO. | TRANS. DATE | ENTRY DATE | DESCRIPTION | DEBIT | CREDIT | DOCUMENT NO.

DEBIT & CREDIT diambil dari kolom "AMOUNT IN BASE CCY".

Output dipecah jadi 2 file:
    1. THC-<MMYYYY awal>-<MMYYYY akhir>.xlsx        -> semua transaksi
    2. THC-NA-<MMYYYY awal>-<MMYYYY akhir>.xlsx      -> transaksi dengan DOCUMENT NO. kosong

Jalankan dengan:
    streamlit run app.py
"""

import io
import os
import re
import shutil
import subprocess
import tempfile
import datetime

import pandas as pd
import requests
import streamlit as st

# --------------------------------------------------------------------------------------
# KONFIGURASI HALAMAN
# --------------------------------------------------------------------------------------
st.set_page_config(page_title="THC Data Merger", page_icon="🔀", layout="wide")

OUTPUT_COLUMNS = [
    "VOUCHER NO.",
    "TRANS. DATE",
    "ENTRY DATE",
    "DESCRIPTION",
    "DEBIT",
    "CREDIT",
    "DOCUMENT NO.",
]

SKIP_PREFIXES = ("ACCOUNT NO", "TOTAL", "ENDING BALANCE", "GRAND TOTAL", "BEGINNING BALANCE")


# ========================================================================================
# CORE PARSING LOGIC
# ========================================================================================

def normalize(cell):
    """Normalisasi nilai sel mentah jadi string bersih (tahan non-breaking space)."""
    if cell is None:
        return ""
    if isinstance(cell, float) and pd.isna(cell):
        return ""
    try:
        if pd.isna(cell):
            return ""
    except (TypeError, ValueError):
        pass
    return str(cell).replace("\xa0", " ").strip()


def read_raw(file_obj, filename):
    """Baca file input (xlsb/xls/xlsx/csv) jadi DataFrame mentah tanpa header."""
    name = filename.lower()
    if name.endswith(".xlsb"):
        return pd.read_excel(file_obj, engine="pyxlsb", header=None, dtype=object)
    elif name.endswith(".xls"):
        return pd.read_excel(file_obj, engine="xlrd", header=None, dtype=object)
    elif name.endswith(".xlsx") or name.endswith(".xlsm"):
        return pd.read_excel(file_obj, engine="openpyxl", header=None, dtype=object)
    elif name.endswith(".csv") or name.endswith(".txt"):
        try:
            return pd.read_csv(file_obj, header=None, sep=None, engine="python", dtype=object)
        except Exception:
            file_obj.seek(0)
            return pd.read_csv(file_obj, header=None, sep=";", engine="python", dtype=object)
    else:
        raise ValueError(f"Format file tidak didukung: {filename}")


def parse_date_cell(val):
    """Ubah nilai tanggal (serial excel / string / datetime) jadi pandas Timestamp."""
    if val is None:
        return pd.NaT
    if isinstance(val, float) and pd.isna(val):
        return pd.NaT
    try:
        if pd.isna(val):
            return pd.NaT
    except (TypeError, ValueError):
        pass

    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return pd.to_datetime(val)

    if isinstance(val, (int, float)):
        try:
            return pd.to_datetime(val, unit="D", origin="1899-12-30")
        except Exception:
            return pd.NaT

    s = str(val).strip()
    if s == "":
        return pd.NaT

    if re.fullmatch(r"\d+(\.\d+)?", s):
        try:
            return pd.to_datetime(float(s), unit="D", origin="1899-12-30")
        except Exception:
            pass

    for fmt in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y"):
        try:
            return pd.to_datetime(s, format=fmt)
        except Exception:
            continue

    try:
        return pd.to_datetime(s, dayfirst=True, errors="coerce")
    except Exception:
        return pd.NaT


def to_number(val):
    """Ubah nilai debit/credit jadi float, default 0.0 kalau kosong/tidak valid."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        try:
            if pd.isna(val):
                return 0.0
        except (TypeError, ValueError):
            pass
        return float(val)
    s = str(val).replace("\xa0", " ").strip()
    if s == "":
        return 0.0
    s = s.replace(",", "").replace(" ", "")
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    try:
        num = float(s)
    except ValueError:
        return 0.0
    return -num if neg else num


def extract_transactions(raw_df, source_name):
    """
    Scan DataFrame mentah (tanpa header) baris per baris, cari blok header
    (VOUCHER NO. / TRANS. DATE / ... / DOCUMENT NO.) lalu ambil setiap baris
    transaksi di bawahnya sampai ketemu baris TOTAL / ENDING BALANCE / GRAND TOTAL
    / Account No. berikutnya (mendukung banyak blok akun & multi-halaman dalam 1 file).
    """
    records = []
    n_rows, n_cols = raw_df.shape
    col_map = None

    row_idx = 0
    while row_idx < n_rows:
        row = [normalize(raw_df.iat[row_idx, c]) for c in range(n_cols)]
        upper_row = [c.upper() for c in row]

        if "VOUCHER NO." in upper_row and "TRANS. DATE" in upper_row:
            col_map = {
                "voucher": upper_row.index("VOUCHER NO."),
                "transdate": upper_row.index("TRANS. DATE"),
                "entrydate": upper_row.index("ENTRY DATE") if "ENTRY DATE" in upper_row else None,
                "desc": upper_row.index("DESCRIPTION") if "DESCRIPTION" in upper_row else None,
            }
            amt_idx = next((i for i, v in enumerate(upper_row) if v.startswith("AMOUNT IN BASE CCY")), None)
            col_map["debit"] = amt_idx
            col_map["credit"] = amt_idx + 1 if amt_idx is not None else None
            doc_idx = next((i for i, v in enumerate(upper_row) if v.startswith("DOCUMENT NO")), None)
            col_map["document"] = doc_idx

            row_idx += 2  # lewati baris header utama + baris sub-header DEBIT/CREDIT
            continue

        if col_map is None:
            row_idx += 1
            continue

        voucher_val = row[col_map["voucher"]] if col_map["voucher"] is not None else ""
        voucher_upper = voucher_val.upper()
        is_skip_row = voucher_val == "" or any(voucher_upper.startswith(p) for p in SKIP_PREFIXES)

        if not is_skip_row:
            transdate_raw = raw_df.iat[row_idx, col_map["transdate"]] if col_map["transdate"] is not None else None
            transdate = parse_date_cell(transdate_raw)
            if pd.notna(transdate):
                entrydate_raw = raw_df.iat[row_idx, col_map["entrydate"]] if col_map["entrydate"] is not None else None
                desc_val = row[col_map["desc"]] if col_map["desc"] is not None else ""
                debit_raw = raw_df.iat[row_idx, col_map["debit"]] if col_map["debit"] is not None else None
                credit_raw = raw_df.iat[row_idx, col_map["credit"]] if col_map["credit"] is not None else None
                doc_val = row[col_map["document"]] if col_map["document"] is not None else ""

                records.append({
                    "VOUCHER NO.": voucher_val,
                    "TRANS. DATE": transdate,
                    "ENTRY DATE": parse_date_cell(entrydate_raw),
                    "DESCRIPTION": desc_val,
                    "DEBIT": to_number(debit_raw),
                    "CREDIT": to_number(credit_raw),
                    "DOCUMENT NO.": doc_val,
                    "_SOURCE FILE": source_name,
                })

        row_idx += 1

    return records


def mmyyyy(year, month):
    return f"{month:02d}{year:04d}"


# ========================================================================================
# EXPORT HELPERS
# ========================================================================================

def prep_for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Rapikan tipe data sebelum ditulis ke file (tanggal jadi date murni, tanpa jam)."""
    out = df.copy()
    for col in ("TRANS. DATE", "ENTRY DATE"):
        out[col] = pd.to_datetime(out[col], errors="coerce").dt.date
    return out


def to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    out = prep_for_export(df)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="THC")
        ws = writer.sheets["THC"]
        widths = [22, 14, 14, 40, 16, 16, 26]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = w
    return buf.getvalue()


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    out = prep_for_export(df)
    buf = io.StringIO()
    out.to_csv(buf, index=False, date_format="%d/%m/%Y")
    return buf.getvalue().encode("utf-8-sig")


def _soffice_binary():
    return shutil.which("soffice") or shutil.which("libreoffice")


# ----------------------------------------------------------------------------------------
# Export ke .xlsb lewat Excel Converter API sendiri (SheetJS backend -> punya writer xlsb asli)
# ----------------------------------------------------------------------------------------
EXCEL_API_BASE = "https://affogateo-excelconverter.hf.space"


def convert_via_excel_api(xlsx_bytes: bytes, base_filename: str, target_format: str = "xlsb", timeout: int = 90) -> bytes:
    """
    Panggil backend Excel Converter API (POST /api/convert/excel lalu GET /api/download/...).
    Raise Exception dengan pesan jelas kalau gagal (space bisa lagi 'sleeping' -> perlu waktu bangun).
    """
    files = {
        "file": (
            f"{base_filename}.xlsx",
            xlsx_bytes,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    }
    data = {"targetFormat": target_format}

    resp = requests.post(f"{EXCEL_API_BASE}/api/convert/excel", files=files, data=data, timeout=timeout)

    try:
        payload = resp.json()
    except ValueError:
        resp.raise_for_status()
        raise RuntimeError("Respons API bukan JSON yang valid.")

    if resp.status_code != 200 or payload.get("error"):
        raise RuntimeError(payload.get("detail", f"HTTP {resp.status_code} dari API konversi."))

    conversion_id = payload.get("conversion_id")
    out_filename = payload.get("filename")
    if not conversion_id or not out_filename:
        raise RuntimeError(f"Respons API tidak lengkap: {payload}")

    dl = requests.get(f"{EXCEL_API_BASE}/api/download/{conversion_id}/{out_filename}", timeout=timeout)
    dl.raise_for_status()
    return dl.content


def to_xlsb_bytes_local(df: pd.DataFrame):
    """
    (Fallback lokal, jarang berhasil) Konversi lewat LibreOffice headless.
    Kebanyakan instalasi LibreOffice bisa MEMBACA .xlsb tapi tidak punya filter untuk MENULISNYA.
    """
    binary = _soffice_binary()
    if binary is None:
        return None

    xlsx_bytes = to_xlsx_bytes(df)
    with tempfile.TemporaryDirectory() as tmpdir:
        xlsx_path = os.path.join(tmpdir, "data.xlsx")
        with open(xlsx_path, "wb") as f:
            f.write(xlsx_bytes)
        try:
            result = subprocess.run(
                [binary, "--headless", "--norestore", "--convert-to", "xlsb", "--outdir", tmpdir, xlsx_path],
                capture_output=True, timeout=120,
            )
        except Exception:
            return None
        xlsb_path = os.path.join(tmpdir, "data.xlsb")
        if result.returncode != 0 or not os.path.exists(xlsb_path):
            return None
        with open(xlsb_path, "rb") as f:
            return f.read()


# ========================================================================================
# UI
# ========================================================================================

st.title("🔀 THC Data Merger")
st.caption("Gabungkan beberapa file transaksi THC bulanan (mentah/data awal) jadi satu file rapi.")

with st.expander("📖 Cara pakai (tutorial)", expanded=True):
    st.markdown(
        """
1. **Download file** `.xlsb` atau tarikan langsung dari **MDIS** (biasanya bentuk `.xls`) — bisa juga `.xlsx` atau `.csv`.
2. **Upload beberapa file** sekaligus (per bulan) di kotak upload di bawah, lalu klik **Proses & Gabungkan**.
3. **Hasil output ada 2 file (saling terpisah, tidak dobel):**
   - Transaksi yang **Document No.-nya terisi** — `THC-MMYYYY-MMYYYY.xlsx`
   - Transaksi yang **Document No.-nya kosong (N/A)** — `THC-NA-MMYYYY-MMYYYY.xlsx`

   Contoh: `THC-012026-082026.xlsx` (data dari Januari 2026 s/d Agustus 2026).

Header hasil akhir: `VOUCHER NO. | TRANS. DATE | ENTRY DATE | DESCRIPTION | DEBIT | CREDIT | DOCUMENT NO.`
(DEBIT & CREDIT diambil dari kolom **AMOUNT IN BASE CCY**).

Tujuan: mempermudah merge beberapa file THC (data mentah/awal) sebelum diolah lebih lanjut.
        """
    )

uploaded_files = st.file_uploader(
    "Upload file THC bulanan (bisa pilih banyak file sekaligus)",
    type=["xlsb", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

process_clicked = st.button("🚀 Proses & Gabungkan", type="primary", disabled=not uploaded_files)

if process_clicked and uploaded_files:
    all_records = []
    errors = []
    progress = st.progress(0.0, text="Memproses file...")

    for i, f in enumerate(uploaded_files):
        try:
            raw = read_raw(f, f.name)
            recs = extract_transactions(raw, f.name)
            if len(recs) == 0:
                errors.append(f"⚠️ **{f.name}**: tidak ada baris transaksi yang terdeteksi (cek format file).")
            all_records.extend(recs)
        except Exception as e:
            errors.append(f"❌ **{f.name}**: gagal diproses — {e}")
        progress.progress((i + 1) / len(uploaded_files), text=f"Memproses {f.name} ({i+1}/{len(uploaded_files)})")

    progress.empty()

    for err in errors:
        st.warning(err) if err.startswith("⚠️") else st.error(err)

    if not all_records:
        st.error("Tidak ada data transaksi yang berhasil diambil dari file yang diupload.")
    else:
        full_df = pd.DataFrame(all_records)

        doc_clean = full_df["DOCUMENT NO."].astype(str).str.strip()
        with_doc_df = full_df[doc_clean != ""].reset_index(drop=True)
        na_doc_df = full_df[doc_clean == ""].reset_index(drop=True)

        periods = sorted({(d.year, d.month) for d in full_df["TRANS. DATE"] if pd.notna(d)})
        if periods:
            start_label = mmyyyy(*periods[0])
            end_label = mmyyyy(*periods[-1])
        else:
            start_label = end_label = "000000"

        merged_filename_base = f"THC-{start_label}-{end_label}"
        na_filename_base = f"THC-NA-{start_label}-{end_label}"

        st.session_state["merged_df"] = with_doc_df[OUTPUT_COLUMNS]
        st.session_state["na_df"] = na_doc_df[OUTPUT_COLUMNS]
        st.session_state["merged_name"] = merged_filename_base
        st.session_state["na_name"] = na_filename_base
        st.session_state["stats"] = {
            "total_files": len(uploaded_files),
            "total_rows": len(full_df),
            "with_doc": len(with_doc_df),
            "na_doc": len(na_doc_df),
            "debit_sum": full_df["DEBIT"].sum(),
            "credit_sum": full_df["CREDIT"].sum(),
            "period": f"{periods[0][1]:02d}/{periods[0][0]} – {periods[-1][1]:02d}/{periods[-1][0]}" if periods else "-",
        }
        st.success("✅ Berhasil digabungkan!")

if "merged_df" in st.session_state:
    stats = st.session_state["stats"]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("File diproses", stats["total_files"])
    c2.metric("Total baris", f'{stats["total_rows"]:,}')
    c3.metric("Document No. kosong", f'{stats["na_doc"]:,}')
    c4.metric("Periode", stats["period"])

    tab1, tab2 = st.tabs(["📄 Document No. Terisi", "🚫 Document No. Kosong (N/A)"])

    with tab1:
        st.dataframe(st.session_state["merged_df"], use_container_width=True, height=350)
    with tab2:
        st.dataframe(st.session_state["na_df"], use_container_width=True, height=350)

    st.divider()
    st.subheader("⬇️ Download")
    st.caption(
        "Export **.xlsb** dilakukan lewat Excel Converter API. Kalau baru pertama kali dipakai "
        "(space lagi 'tidur'), proses convert bisa makan waktu sampai ~1 menit untuk 'membangunkan' server-nya."
    )

    def download_section(label, df, base_name):
        st.markdown(f"**{label}** — `{base_name}.xlsx`")
        col_x, col_c, col_b = st.columns(3)
        with col_x:
            st.download_button(
                "📗 Download .xlsx",
                data=to_xlsx_bytes(df),
                file_name=f"{base_name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"xlsx_{base_name}",
            )
        with col_c:
            st.download_button(
                "📄 Download .csv",
                data=to_csv_bytes(df),
                file_name=f"{base_name}.csv",
                mime="text/csv",
                use_container_width=True,
                key=f"csv_{base_name}",
            )
        with col_b:
            xlsb_state_key = f"xlsb_bytes_{base_name}"
            if xlsb_state_key in st.session_state:
                st.download_button(
                    "📘 Download .xlsb",
                    data=st.session_state[xlsb_state_key],
                    file_name=f"{base_name}.xlsb",
                    mime="application/vnd.ms-excel.sheet.binary.macroenabled.12",
                    use_container_width=True,
                    key=f"xlsb_dl_{base_name}",
                )
                if st.button("🔄 Convert ulang", key=f"xlsb_redo_{base_name}", use_container_width=True):
                    del st.session_state[xlsb_state_key]
                    st.rerun()
            else:
                if st.button("📘 Convert ke .xlsb", key=f"xlsb_convert_{base_name}", use_container_width=True):
                    try:
                        with st.spinner("Menghubungi server konversi... (bisa sampai ~1 menit kalau server baru bangun)"):
                            xlsx_bytes = to_xlsx_bytes(df)
                            xlsb_bytes = convert_via_excel_api(xlsx_bytes, base_name, target_format="xlsb")
                        st.session_state[xlsb_state_key] = xlsb_bytes
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal convert ke .xlsb: {e}")
                        with st.expander("Detail error (buat debug)"):
                            st.code(str(e))
                        st.info(
                            "Alternatif: download .xlsx dulu, lalu di Excel: "
                            "**File → Save As → Excel Binary Workbook (*.xlsb)**."
                        )

    download_section("Document No. Terisi", st.session_state["merged_df"], st.session_state["merged_name"])
    st.write("")
    download_section("Document No. Kosong (N/A)", st.session_state["na_df"], st.session_state["na_name"])

else:
    st.info("Upload file lalu klik **Proses & Gabungkan** untuk mulai.")
